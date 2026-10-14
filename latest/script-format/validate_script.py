#!/usr/bin/env python3
"""Reference validator for DPM-XL Script Documents.

Checks a Script Document against the structural validation rules V-1 .. V-32 of
the Script Format specification, chapter 5.3. Uses only the Python standard
library; no DPM database or taxonomy is required, because every rule is
checkable from the document alone.

Usage:
    python validate_script.py FILE [FILE ...]
    python validate_script.py --json FILE
    python validate_script.py --quiet FILE      # exit code only

Exit codes:
    0   every document is valid
    1   at least one document violates at least one rule
    2   a document could not be read or parsed
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

# --------------------------------------------------------------------------
# Tables from the specification
# --------------------------------------------------------------------------

#: §3.2 — the fourteen REQUIRED members of the Module Object.
MODULE_MEMBERS = frozenset({
    "format_version",
    "module_code", "module_version", "framework_code", "dpm_release", "dates",
    "operations", "variables", "tables", "preconditions",
    "precondition_variables", "parameters", "dependency_information",
    "dependency_modules",
})

#: §3.10 — the members of an Operation Object.
OPERATION_MEMBERS = frozenset({
    "code", "version_id", "expression", "root_operator_id", "ast",
    "severity", "from_submission_date",
})

#: §3.7 — the members of a precondition entry.
PRECONDITION_MEMBERS = frozenset({"code", "version_id", "ast",
                                  "affected_operations"})

#: §3.12 — the members of a dependency_modules entry.
DEPENDENCY_MODULE_MEMBERS = frozenset({"tables", "variables"})

#: §3.9.1 — the twelve Module Parameter type names: the surface type words of
#: Information Model §3.1.1, each word capitalised.
SCALAR_PARAMETER_TYPES = frozenset({"Number", "Integer", "String", "Date",
                                    "Boolean", "Item"})
PARAMETER_TYPES = SCALAR_PARAMETER_TYPES | frozenset(
    "Set" + name for name in SCALAR_PARAMETER_TYPES)

#: §4.8.1 — the JSON type a plain Scalar `param_type` requires of `default`.
#: "Item" and each element of a Set default are checked separately, see check_parameter_ref().
PARAMETER_DEFAULT_TYPES = {"Integer": int, "Number": float, "Boolean": bool,
                           "String": str, "Date": str}

#: §3.10 — permitted severities.
SEVERITIES = frozenset({"error", "warning", "info"})

#: §3.5 — permitted Data Type codes (DataType.Code in DPM-ML).
DATA_TYPES = frozenset({"i", "r", "s", "es", "b", "t", "e", "m", "p", "d",
                        "dt", "u", "o"})

#: §4.1.2 — Node classes and their permitted members, excluding `class_name`.
NODE_MEMBERS = {
    "VarID": frozenset({"table", "row", "column", "sheet", "interval",
                        "default", "data"}),
    "Constant": frozenset({"type_", "value"}),
    "Scalar": frozenset({"item", "scalar_type"}),
    "Set": frozenset({"children"}),
    "Dimension": frozenset({"dimension_code"}),
    "PreconditionItem": frozenset({"variable_id", "variable_code"}),
    "ParExpr": frozenset({"expression"}),
    "UnaryOp": frozenset({"op", "operand"}),
    "BinOp": frozenset({"op", "left", "right"}),
    "ComplexNumericOp": frozenset({"op", "operands"}),
    "SetOp": frozenset({"op", "operands"}),
    "SubstrOp": frozenset({"op", "operand", "start", "length"}),
    "DateConstructorOp": frozenset({"op", "year_node", "month_node", "day_node"}),
    "AnnualiseOp": frozenset({"op", "operand", "fy_end", "component"}),
    "AggregationOp": frozenset({"op", "operand", "grouping_clause",
                                "analytic_clause"}),
    "GroupingClause": frozenset({"components"}),
    "AnalyticClause": frozenset({"partition_by", "order_by", "window"}),
    "WhereClauseOp": frozenset({"operand", "condition"}),
    "SubClauseOp": frozenset({"operand", "condition"}),
    "GetClauseOp": frozenset({"operand", "component"}),
    "RenameClauseOp": frozenset({"operand", "clauses"}),
    "FilterOp": frozenset({"selection", "condition"}),
    "CondExpr": frozenset({"condition", "then_expr", "else_expr"}),
    "ParameterRef": frozenset({"code", "param_type", "default"}),
    "TimeShiftOp": frozenset({"operand", "period_indicator", "shift_number",
                              "reference_period"}),
}

#: §4.1.3 — members whose absence is significant, so V-5 must not require them.
OPTIONAL_MEMBERS = {
    "VarID": frozenset({"row", "column", "sheet", "default"}),
    "CondExpr": frozenset({"else_expr"}),
    "SubstrOp": frozenset({"start", "length"}),
}

#: §4.11, §4.12, §4.14, §4.18 — permitted `op` values per class.
#: §4.11 — the six date-component extractors share one Node shape.
DATE_COMPONENTS = frozenset({"year", "semester", "quarter", "month", "week",
                             "day"})

OPS = {
    "UnaryOp": frozenset({"+", "-", "abs", "sqrt", "exp", "ln", "isnull",
                          "not", "len"}) | DATE_COMPONENTS,
    "BinOp": frozenset({"+", "-", "*", "/", "power", "log", "nvl", "&",
                        "=", "!=", "<", "<=", ">", ">=", "match", "in",
                        "and", "or", "xor"}),
    "ComplexNumericOp": frozenset({"min", "max"}),
    "SetOp": frozenset({"set_of", "union", "intersect", "setdiff", "symdiff"}),
    "SubstrOp": frozenset({"substr"}),
    "DateConstructorOp": frozenset({"date"}),
    "AnnualiseOp": frozenset({"annualise"}),
    "AggregationOp": frozenset({"sum", "count", "avg", "median", "max_aggr",
                                "min_aggr", "rank"}),
}

#: §4.14 — (minimum, maximum) arity of each SetOp `op`; None means unbounded.
SETOP_ARITY = {"set_of": (1, 1), "union": (2, None), "intersect": (2, None),
               "setdiff": (2, 2), "symdiff": (2, 2)}

#: §4.20 — permitted window frame types and bound kinds.
WINDOW_FRAME_TYPES = frozenset({"data_points", "range"})
WINDOW_BOUND_TYPES = frozenset({"unbounded_preceding", "n_preceding",
                                "current_data_point", "n_following",
                                "unbounded_following"})
#: Window order, unbounded_preceding (most preceding) to unbounded_following.
WINDOW_BOUND_ORDER = {"unbounded_preceding": 0, "n_preceding": 1,
                      "current_data_point": 2, "n_following": 3,
                      "unbounded_following": 4}
WINDOW_N_BOUND_TYPES = frozenset({"n_preceding", "n_following"})

#: §4.1.2 — Nodes denoting data or literals rather than Operator applications.
OPERAND_CLASSES = frozenset({"VarID", "Constant", "Scalar", "Set", "Dimension",
                             "ParameterRef", "PreconditionItem"})

#: §4.2 — (class_name, op) -> OperatorID. `op` is None for single-Operator classes.
OPERATOR_IDS = {
    ("UnaryOp", "+"): 1, ("BinOp", "+"): 2, ("BinOp", "/"): 3,
    ("UnaryOp", "-"): 4, ("BinOp", "-"): 5, ("UnaryOp", "abs"): 6,
    ("ComplexNumericOp", "min"): 7, ("BinOp", "*"): 8,
    ("ComplexNumericOp", "max"): 9, ("BinOp", "="): 13, ("BinOp", "<="): 14,
    ("BinOp", ">="): 15, ("BinOp", "in"): 16, ("UnaryOp", "isnull"): 17,
    ("BinOp", ">"): 18, ("BinOp", "<"): 19, ("BinOp", "!="): 20,
    ("BinOp", "and"): 22, ("BinOp", "or"): 23, ("UnaryOp", "not"): 24,
    ("UnaryOp", "sqrt"): 10, ("BinOp", "match"): 21, ("BinOp", "xor"): 25,
    ("AggregationOp", "max_aggr"): 11, ("AggregationOp", "min_aggr"): 12,
    ("AggregationOp", "sum"): 26, ("AggregationOp", "count"): 27,
    ("WhereClauseOp", None): 28, ("GetClauseOp", None): 29,
    ("CondExpr", None): 30, ("FilterOp", None): 31, ("TimeShiftOp", None): 32,
    ("RenameClauseOp", None): 33, ("GroupingClause", None): 35,
    ("ParExpr", None): 37, ("SubClauseOp", None): 38,
}

#: §4.4 — Constant.type_ -> permitted JSON type of Constant.value.
CONSTANT_TYPES = {"Integer": int, "Number": float, "Boolean": bool,
                  "String": str, "Date": str, "Null": type(None)}

#: §3.11 — permitted ref_period forms: T, T-<n><p> and T+<n><p>.
REF_PERIOD = re.compile(r"^T$|^T[+-]\d+[A-Z]$")
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

#: §4.7, §4.23 — codes that name no Open Key and need no declaration.
RESERVED_COMPONENTS = frozenset({"f", "entityID", "baseCurrency", "refPeriod"})

#: §4.1.5 — sentinel for an index set whose cardinality is data-dependent and
#: therefore not statically determinable (a grouped aggregation).
OPAQUE = "OPAQUE"


class Report:
    """Collects violations keyed by rule identifier."""

    def __init__(self) -> None:
        self.violations: dict[str, list[str]] = defaultdict(list)
        self.unverified = 0

    def fail(self, rule: str, detail: str) -> None:
        self.violations[rule].append(detail)

    @property
    def count(self) -> int:
        return sum(len(v) for v in self.violations.values())

    @property
    def valid(self) -> bool:
        return self.count == 0

    def rules(self) -> list[str]:
        return sorted(self.violations, key=lambda r: int(r.split("-")[1]))


# --------------------------------------------------------------------------
# Tree traversal
# --------------------------------------------------------------------------

def child_nodes(node):
    """Yield (member_name, child_node) for each Node directly beneath `node`."""
    for key, value in node.items():
        if key == "class_name":
            continue
        if isinstance(value, dict) and "class_name" in value:
            yield key, value
        elif isinstance(value, list):
            for element in value:
                if isinstance(element, dict) and "class_name" in element:
                    yield key, element


def walk(node, parent=None, member=None):
    """Yield (node, parent, member_name) for `node` and every descendant."""
    yield node, parent, member
    for key, child in child_nodes(node):
        yield from walk(child, node, key)


def axes_of(index_set_):
    """§4.1.5 — the axes on which the indices of a set carry a component."""
    present = set()
    for position in index_set_:
        for slot, axis in enumerate("xyz"):
            if position[slot] is not None:
                present.add(axis)
    return frozenset(present)


def project(index_set_, axes):
    """§4.1.5 — the projection of an index set onto `axes`."""
    slots = [slot for slot, axis in enumerate("xyz") if axis in axes]
    return frozenset(tuple(position[slot] for slot in slots)
                     for position in index_set_)


def index_set(node, report: Report, where: str):
    """Return the Position Index set of `node` per §4.1.5, or OPAQUE.

    Reports a V-17 violation for any ill-formed Operator encountered.
    """
    cls = node["class_name"]

    if cls == "VarID":
        return frozenset((d.get("x"), d.get("y"), d.get("z"))
                         for d in node["data"])
    if cls in OPERAND_CLASSES:
        return frozenset()
    if cls == "ParExpr":
        return index_set(node["expression"], report, where)
    if cls == "UnaryOp":
        return index_set(node["operand"], report, where)
    if cls in ("WhereClauseOp", "SubClauseOp", "TimeShiftOp",
               "RenameClauseOp", "GetClauseOp", "SubstrOp", "AnnualiseOp"):
        return index_set(node["operand"], report, where)
    if cls == "FilterOp":
        return index_set(node["selection"], report, where)
    if cls == "SetOp":
        # §4.1.5 — the result is a Scalar Set, which carries no positions.
        for operand in node["operands"]:
            index_set(operand, report, where)
        return frozenset()
    if cls == "AggregationOp":
        inner = index_set(node["operand"], report, where)
        if node.get("analytic_clause") is not None:
            return inner                # analytic: one value per input Record
        if node.get("grouping_clause") is None:
            return frozenset()          # reduced to a Scalar
        return OPAQUE                   # one index per group, data-dependent

    # A malformed Node must be reported by V-5, not crash the walk, so every
    # operand is fetched defensively and missing ones are skipped.
    if cls == "DateConstructorOp":
        operands = [node.get("year_node"), node.get("month_node"),
                    node.get("day_node")]
    elif cls == "ComplexNumericOp":
        operands = node.get("operands") or []
    elif cls == "BinOp":
        operands = [node.get("left"), node.get("right")]
    elif cls == "CondExpr":
        operands = [node[k] for k in ("condition", "then_expr", "else_expr")
                    if node.get(k) is not None]
    elif cls == "ParameterRef":
        return frozenset()
    else:
        return frozenset()

    operands = [o for o in operands
                if isinstance(o, dict) and "class_name" in o]
    if not operands:
        return frozenset()

    sets = [index_set(o, report, where) for o in operands]
    if any(s is OPAQUE for s in sets):
        report.unverified += 1
        return OPAQUE

    # §4.1.5: S is the set of non-empty operand index sets. Operands combine by
    # natural join on the axes they share, so it is the axes that must line up,
    # not the arity.
    distinct = {s for s in sets if s}
    if not distinct:                                        # case 1: Scalar
        return frozenset()

    widest = max(distinct, key=lambda s: len(axes_of(s)))
    for operand_set in distinct:
        axes = axes_of(operand_set)
        if not axes <= axes_of(widest):                     # case 5
            report.fail("V-17", f"{where}: operand axes "
                                f"{sorted(axes)} and {sorted(axes_of(widest))} "
                                f"have no common widest operand")
            return OPAQUE
        if project(widest, axes) != project(operand_set, axes):
            report.fail("V-17", f"{where}: operands disagree on axes "
                                f"{sorted(axes) or ['(none)']} — "
                                f"{len(project(operand_set, axes))} position(s) "
                                f"against {len(project(widest, axes))}")
            return OPAQUE
    return widest

# --------------------------------------------------------------------------
# Rule checks
# --------------------------------------------------------------------------

def check_document(raw: dict, report: Report) -> None:
    # V-1
    if len(raw) != 1:
        report.fail("V-1", f"root object has {len(raw)} members, expected 1")
        return
    uri, module = next(iter(raw.items()))
    if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://\S+$", uri):
        report.fail("V-1", f"root member name is not an absolute URI: {uri!r}")

    # V-2
    if set(module) != MODULE_MEMBERS:
        missing = MODULE_MEMBERS - set(module)
        extra = set(module) - MODULE_MEMBERS
        if missing:
            report.fail("V-2", f"missing member(s): {sorted(missing)}")
        if extra:
            report.fail("V-2", f"unrecognised member(s): {sorted(extra)}")
        if missing:
            return

    # V-2 — §3.2.1. MAJOR.MINOR, as published.
    if not re.match(r"^\d+\.\d+$", str(module["format_version"])):
        report.fail("V-2", f"format_version {module['format_version']!r} "
                           f"is not MAJOR.MINOR")

    tables = module["tables"]
    dep_modules = module["dependency_modules"]
    operations = module["operations"]
    parameters = module["parameters"]

    # V-29 — the declared type names (§3.9.1).
    for name, declared in sorted(parameters.items()):
        if declared not in PARAMETER_TYPES:
            report.fail("V-29", f"parameters/{name}: type {declared!r} is not "
                                f"one of the twelve names of §3.9.1")

    # V-32 — a dependency_modules entry carries exactly `tables` and `variables`.
    for dep_uri, dep in sorted(dep_modules.items()):
        if not isinstance(dep, dict) or set(dep) != DEPENDENCY_MODULE_MEMBERS:
            report.fail("V-32", f"dependency_modules/{dep_uri}: members "
                                f"{sorted(dep) if isinstance(dep, dict) else dep!r}")

    # Data Type lookup, per §3.5, §3.6 and §3.11.
    local_types: dict[str, str] = dict(module["variables"])
    foreign_types: dict[str, str] = {}
    foreign_tables: dict[str, str] = {}
    for dep_uri, dep in dep_modules.items():
        if not isinstance(dep, dict):
            continue
        foreign_types.update(dep.get("variables") or {})
        for table_code, table in (dep.get("tables") or {}).items():
            foreign_tables[table_code] = dep_uri
            foreign_types.update(table["variables"])

    def open_keys(table_code: str) -> set[str]:
        if table_code in tables:
            return set(tables[table_code]["open_keys"])
        owner = foreign_tables.get(table_code)
        if owner:
            return set(dep_modules[owner]["tables"][table_code]["open_keys"])
        return set()

    #: §4.3.4 — operand_reference_id -> the one cell it may denote.
    cell_of_ref: dict = {}

    check_dictionaries(module, tables, report)
    check_dates(module, report)

    indicators = module["precondition_variables"]
    for code, operation in sorted(operations.items()):
        check_operation(code, operation, tables, foreign_tables, local_types,
                        foreign_types, open_keys, parameters, indicators,
                        cell_of_ref, report)

    check_preconditions(module, operations, tables, foreign_tables, local_types,
                        foreign_types, open_keys, parameters, cell_of_ref,
                        report)
    check_dependencies(module, uri, operations, dep_modules, report)


def check_dictionaries(module, tables, report: Report) -> None:
    # V-8
    union: set[str] = set()
    for table_code, table in tables.items():
        union |= set(table["variables"])
        for var_id, code in table["variables"].items():
            if code not in DATA_TYPES:
                report.fail("V-8", f"tables/{table_code}/{var_id}: "
                                   f"unknown Data Type code {code!r}")
    declared = set(module["variables"])
    if declared != union:
        report.fail("V-8", f"`variables` has {len(declared)} keys, union of "
                           f"`tables[*].variables` has {len(union)}; "
                           f"symmetric difference {len(declared ^ union)}")
    for var_id, code in module["variables"].items():
        if code not in DATA_TYPES:
            report.fail("V-8", f"variables/{var_id}: unknown code {code!r}")

    # V-26 — the Data Type must be a function of the VariableID alone, so that a
    # consumer loading XBRL-CSV facts can type-check on the VariableID before it
    # knows which table or Operation will consume the fact (§3.5).
    assignments: dict[str, set[str]] = defaultdict(set)
    for table in tables.values():
        for var_id, code in table["variables"].items():
            assignments[var_id].add(code)
    for var_id, code in module["variables"].items():
        assignments[var_id].add(code)
    for dependency in module["dependency_modules"].values():
        for var_id, code in dependency["variables"].items():
            assignments[var_id].add(code)
        for table in dependency["tables"].values():
            for var_id, code in table["variables"].items():
                assignments[var_id].add(code)
    for var_id, codes in sorted(assignments.items()):
        if len(codes) > 1:
            report.fail("V-26", f"Variable {var_id} is assigned Data Types "
                                f"{sorted(codes)}")

    # V-11
    overlap = set(module["precondition_variables"]) & declared
    if overlap:
        report.fail("V-11", f"{len(overlap)} identifier(s) in both "
                            f"`variables` and `precondition_variables`")
    for var_id, code in module["precondition_variables"].items():
        if code != "b":
            report.fail("V-11", f"precondition_variables/{var_id}: {code!r}, "
                                f"expected 'b'")


def check_dates(module, report: Report) -> None:
    # V-18
    dates = module["dates"]
    if not ISO_DATE.match(dates["from"] or ""):
        report.fail("V-18", f"dates.from {dates['from']!r}")
    if dates["to"] is not None:
        if not ISO_DATE.match(dates["to"]):
            report.fail("V-18", f"dates.to {dates['to']!r}")
        elif dates["to"] < dates["from"]:
            report.fail("V-18", f"dates.to {dates['to']} < from {dates['from']}")
    for code, operation in module["operations"].items():
        if not ISO_DATE.match(operation.get("from_submission_date") or ""):
            report.fail("V-18", f"{code}: from_submission_date "
                                f"{operation.get('from_submission_date')!r}")


def check_operation(code, operation, tables, foreign_tables, local_types,
                    foreign_types, open_keys, parameters, declared_indicators,
                    cell_of_ref, report: Report) -> None:
    # V-3
    if operation.get("code") != code:
        report.fail("V-3", f"operations/{code}: code member is "
                           f"{operation.get('code')!r}")
    if set(operation) != OPERATION_MEMBERS:
        missing = OPERATION_MEMBERS - set(operation)
        extra = set(operation) - OPERATION_MEMBERS
        if missing:
            report.fail("V-3", f"operations/{code}: missing {sorted(missing)}")
        if extra:
            report.fail("V-3", f"operations/{code}: unrecognised "
                               f"{sorted(extra)}")
        if "ast" in missing:
            return
    # V-19
    if operation.get("severity") not in SEVERITIES:
        report.fail("V-19", f"{code}: severity {operation.get('severity')!r}")

    root = operation["ast"]

    # V-25
    if root["class_name"] in OPERAND_CLASSES:
        report.fail("V-25", f"{code}: root Node is an Operand "
                            f"({root['class_name']})")
    # V-4 — only where the root Node encodes a DPM-ML Operator (§4.2.1).
    expected = OPERATOR_IDS.get((root["class_name"], root.get("op")))
    if expected is not None and expected != operation.get("root_operator_id"):
        report.fail("V-4", f"{code}: root_operator_id "
                           f"{operation.get('root_operator_id')}, "
                           f"root Node implies {expected}")

    check_tree(code, root, tables, foreign_tables, local_types,
               foreign_types, open_keys, parameters, declared_indicators,
               cell_of_ref, report)


def check_tree(code, root, tables, foreign_tables, local_types,
               foreign_types, open_keys, parameters, declared_indicators,
               cell_of_ref, report: Report) -> None:
    """Per-Node rules, applied identically to an Operation and a precondition.

    §3.7.1 — a precondition tree is an Operation tree, so every rule that
    governs the Nodes of one governs the Nodes of the other.
    """
    for node, parent, member in walk(root):
        cls = node["class_name"]
        # V-5
        if cls not in NODE_MEMBERS:
            report.fail("V-5", f"{code}: unrecognised class_name {cls!r}")
            continue
        extra = set(node) - {"class_name"} - NODE_MEMBERS[cls]
        if extra:
            report.fail("V-5", f"{code}/{cls}: unrecognised member(s) "
                               f"{sorted(extra)}")
        required = NODE_MEMBERS[cls] - set(node) - OPTIONAL_MEMBERS.get(cls, frozenset())
        if required:
            report.fail("V-5", f"{code}/{cls}: missing member(s) "
                               f"{sorted(required)}")
        if cls in OPS and node.get("op") not in OPS[cls]:
            report.fail("V-5", f"{code}/{cls}: unrecognised op "
                               f"{node.get('op')!r}")

        if cls == "VarID":
            check_varid(code, node, tables, foreign_tables, local_types,
                        foreign_types, cell_of_ref, report)
        elif cls == "Constant":
            check_constant(code, node, report)
        elif cls == "Set":
            check_set(code, node, report)
        elif cls == "BinOp" and node.get("op") == "match":
            # §4.12 constraint 3 — the pattern is a String literal.
            pattern = node.get("right") or {}
            if pattern.get("class_name") != "Constant" \
                    or pattern.get("type_") != "String":
                report.fail("V-31", f"{code}: `match` pattern is "
                                    f"{pattern.get('class_name')}, expected a "
                                    f"String Constant")
        elif cls == "BinOp" and node.get("op") == "in":
            if node["right"].get("class_name") not in ("Set", "SetOp"):
                report.fail("V-21", f"{code}: `in` right operand is "
                                    f"{node['right'].get('class_name')}, "
                                    f"expected Set or SetOp")
        elif cls == "GroupingClause":
            check_grouping(code, node, parent, member, open_keys, report)
        elif cls == "AnalyticClause":
            check_analytic(code, node, parent, member, report)
        elif cls == "AggregationOp":
            # V-30 — the two clauses are mutually exclusive, `rank` is
            # analytic by definition, and `rank` rejects a window clause (§4.18).
            if node.get("grouping_clause") is not None \
                    and node.get("analytic_clause") is not None:
                report.fail("V-30", f"{code}: AggregationOp has both a "
                                    f"grouping_clause and an analytic_clause")
            if node.get("op") == "rank":
                if node.get("analytic_clause") is None:
                    report.fail("V-30", f"{code}: `rank` without an AnalyticClause")
                elif node.get("analytic_clause", {}).get("window") is not None:
                    report.fail("V-30", f"{code}: `rank` cannot have a window clause")
        elif cls == "SetOp":
            check_setop(code, node, report)
        elif cls == "SubstrOp":
            check_substr(code, node, report)
        elif cls == "RenameClauseOp":
            check_rename(code, node, report)
        elif cls == "GetClauseOp":
            if not isinstance(node.get("component"), str):
                report.fail("V-31", f"{code}: GetClauseOp component is "
                                    f"{node.get('component')!r}, expected a String")
        elif cls == "ParameterRef":
            check_parameter_ref(code, node, parameters, report)
        elif cls == "PreconditionItem":
            # V-12 — wherever it occurs, §4.9.
            if str(node.get("variable_id")) not in declared_indicators:
                report.fail("V-12", f"{code}: variable_id "
                                    f"{node.get('variable_id')} is not in "
                                    f"`precondition_variables`")

    check_dimension_placement(code, root, report)
    index_set(root, report, code)


def check_varid(code, node, tables, foreign_tables, local_types,
                foreign_types, cell_of_ref, report: Report) -> None:
    entries = node["data"]

    # V-7
    if not entries:
        report.fail("V-7", f"{code}: VarID with empty `data`")
        return

    # V-6 — §4.3.2. An entry carries the resolved coordinate of every axis the
    # table has, except one already pinned on the VarID, where repeating it is
    # permitted but not required (Producers differ). The index appears exactly
    # where the axis varies; the VarID-level coordinate records only what the
    # selection pinned.
    applicable = False
    for axis, index in (("row", "x"), ("column", "y"), ("sheet", "z")):
        on_node = axis in node
        with_coord = [e for e in entries if axis in e]
        with_index = [e for e in entries if index in e]
        if with_coord and len(with_coord) != len(entries):
            report.fail("V-6", f"{code}: {axis} on {len(with_coord)} of "
                               f"{len(entries)} entries")
            continue
        if with_index and len(with_index) != len(entries):
            report.fail("V-6", f"{code}: {index} on {len(with_index)} of "
                               f"{len(entries)} entries")
            continue
        if not with_coord:
            if with_index:
                report.fail("V-6", f"{code}: {index} present without {axis}")
            elif on_node:
                # Fixed by selection, with the entry coordinate omitted: the
                # axis is pinned one level up and recoverable from there.
                applicable = True
            continue
        applicable = True
        distinct = {e[axis] for e in entries}
        if len(distinct) > 1 and not with_index:
            report.fail("V-6", f"{code}: {axis} takes {len(distinct)} values "
                               f"but no {index} index is present")
        if len(distinct) == 1 and with_index:
            report.fail("V-6", f"{code}: {index} present although {axis} does "
                               f"not vary")
        if on_node:
            if with_index:
                report.fail("V-6", f"{code}: {axis} pinned on the VarID "
                                   f"although it varies")
            elif distinct != {node[axis]}:
                report.fail("V-6", f"{code}: VarID {axis}={node[axis]!r} "
                                   f"disagrees with entries {sorted(distinct)}")
    if not applicable:
        report.fail("V-6", f"{code}: VarID has no applicable axis")

    # V-7, V-15
    seen: set[tuple] = set()
    table_code = node.get("table")
    for entry in entries:
        if entry.get("datapoint") is None:
            report.fail("V-7", f"{code}: null datapoint")
        for index in ("x", "y", "z"):
            value = entry.get(index)
            if index in entry and (not isinstance(value, int)
                                   or isinstance(value, bool) or value < 1):
                report.fail("V-7", f"{code}: {index} is {value!r}, "
                                   f"expected a positive integer")
        position = (entry.get("x"), entry.get("y"), entry.get("z"))
        if position in seen:
            report.fail("V-7", f"{code}: duplicate Position Index {position}")
        seen.add(position)
        # V-15 — an operand_reference_id identifies a resolved cell, so it may
        # recur, but never for two different cells (§4.3.4).
        ref = entry.get("operand_reference_id")
        cell = (table_code, entry.get("datapoint"), entry.get("row"),
                entry.get("column"), entry.get("sheet"))
        previous = cell_of_ref.setdefault(ref, cell)
        if previous != cell:
            report.fail("V-15", f"{code}: operand_reference_id {ref} denotes "
                                f"{cell} here and {previous} elsewhere")

    # V-9, V-10
    table = node["table"]
    if table in tables:
        types = local_types
    elif table in foreign_tables:
        types = foreign_types
    else:
        report.fail("V-10", f"{code}: table {table!r} is in neither `tables` "
                            f"nor `dependency_modules[*].tables`")
        return
    # V-9
    for entry in entries:
        key = str(entry.get("datapoint"))
        if key not in types:
            report.fail("V-9", f"{code}: datapoint {key} of table {table} "
                               f"has no Data Type")


def check_constant(code, node, report: Report) -> None:
    # V-20
    declared = node.get("type_")
    value = node.get("value")
    expected = CONSTANT_TYPES.get(declared)
    if expected is None:
        report.fail("V-20", f"{code}: Constant type_ {declared!r}")
        return
    ok = isinstance(value, expected)
    if declared != "Boolean" and isinstance(value, bool):
        ok = False                      # bool is a subclass of int in Python
    if declared == "Number" and isinstance(value, int) and not isinstance(value, bool):
        ok = True                       # a JSON integer is a valid Number
    if not ok:
        report.fail("V-20", f"{code}: Constant type_ {declared!r} with value "
                            f"of JSON type {type(value).__name__}")


def check_set(code, node, report: Report) -> None:
    # V-21 — `children` may be empty (the `{}` literal, §4.6).
    children = node.get("children")
    if not isinstance(children, list):
        report.fail("V-21", f"{code}: Set `children` is not an array")
        return
    for child in children:
        if child.get("class_name") not in ("Scalar", "Constant"):
            report.fail("V-21", f"{code}: Set child is "
                                f"{child.get('class_name')}, expected Scalar "
                                f"or Constant")


def check_grouping(code, node, parent, member, open_keys, report: Report) -> None:
    # V-22
    if parent is None or parent.get("class_name") != "AggregationOp" \
            or member != "grouping_clause":
        report.fail("V-22", f"{code}: GroupingClause is not the "
                            f"`grouping_clause` of an AggregationOp")
        return
    # V-23
    components = node.get("components") or []
    if not components:
        report.fail("V-23", f"{code}: GroupingClause with empty `components`")
    table = None
    renamed: set[str] = set()
    for descendant, _, _ in walk(parent["operand"]):
        if descendant["class_name"] == "RenameClauseOp":
            renamed.update(c.get("to_component")
                           for c in descendant.get("clauses") or [])
        if descendant["class_name"] == "VarID" and table is None:
            table = descendant["table"]
    # §4.19 — a component resolves against the operand as it presents itself,
    # so a `to_component` introduced by a RenameClauseOp counts.
    keys = (open_keys(table) if table else set()) | renamed | RESERVED_COMPONENTS
    for component in components:
        if component not in ("r", "c", "s") and component not in keys:
            report.fail("V-23", f"{code}: grouping component {component!r} is "
                                f"neither an axis token nor a Component of "
                                f"{table}")


def check_analytic(code, node, parent, member, report: Report) -> None:
    """V-22, V-30 — placement and shape of an AnalyticClause (§4.20)."""
    if parent is None or parent.get("class_name") != "AggregationOp" \
            or member != "analytic_clause":
        report.fail("V-22", f"{code}: AnalyticClause is not the "
                            f"`analytic_clause` of an AggregationOp")
    order_by = node.get("order_by")
    window = node.get("window")
    is_rank = parent is not None and parent.get("op") == "rank"
    must_have_order_by = is_rank or (window is not None)

    if not isinstance(order_by, list):
        report.fail("V-30", f"{code}: AnalyticClause `order_by` is not an array")
    elif must_have_order_by and not order_by:
        report.fail("V-30", f"{code}: AnalyticClause with empty `order_by`")
    elif order_by:
        for key in order_by:
            if not isinstance(key, dict) or "key_name" not in key:
                report.fail("V-30", f"{code}: order_by entry {key!r} has no "
                                    f"`key_name`")
            elif key.get("direction") not in ("asc", "desc"):
                report.fail("V-30", f"{code}: order_by direction "
                                    f"{key.get('direction')!r}")
    if not isinstance(node.get("partition_by"), list):
        report.fail("V-30", f"{code}: `partition_by` is not an array")
    if window is not None:
        check_window(code, window, node.get("order_by"), report)


def check_window(code, window, order_by, report: Report) -> None:
    """V-30 — the frame of a non-null AnalyticClause window (§4.20)."""
    if not isinstance(window, dict) or set(window) != {"frame_type", "start", "end"}:
        report.fail("V-30", f"{code}: window members "
                            f"{sorted(window) if isinstance(window, dict) else window!r}")
        return
    if window["frame_type"] not in WINDOW_FRAME_TYPES:
        report.fail("V-30", f"{code}: window frame_type {window['frame_type']!r}")
    seen = []
    for side in ("start", "end"):
        bound = window[side]
        if not isinstance(bound, dict) or "bound_type" not in bound \
                or bound["bound_type"] not in WINDOW_BOUND_TYPES \
                or set(bound) - {"bound_type", "n"}:
            report.fail("V-30", f"{code}: window {side} {bound!r}")
            return
        kind = bound["bound_type"]
        if kind not in WINDOW_N_BOUND_TYPES and "n" in bound:
            report.fail("V-30", f"{code}: window {side} {kind!r} carries an `n`")
        if kind in WINDOW_N_BOUND_TYPES and (not isinstance(bound.get("n"), int)
                                             or isinstance(bound.get("n"), bool)
                                             or bound["n"] < 1):
            report.fail("V-30", f"{code}: window {side} n {bound.get('n')!r}")
        seen.append(WINDOW_BOUND_ORDER[kind])
    if seen[0] > seen[1]:
        report.fail("V-30", f"{code}: window start is after its end")
    if window["frame_type"] == "range" and any(
            window[s].get("bound_type") in WINDOW_N_BOUND_TYPES
            for s in ("start", "end")) \
            and len(order_by or []) != 1:
        report.fail("V-30", f"{code}: `range` window with an `n` bound needs "
                            f"exactly one order_by key")


def check_setop(code, node, report: Report) -> None:
    """V-31 — SetOp arity (§4.14)."""
    operands = node.get("operands")
    if not isinstance(operands, list):
        report.fail("V-31", f"{code}: SetOp `operands` is not an array")
        return
    bounds = SETOP_ARITY.get(node.get("op"))
    if bounds is None:
        return                          # unrecognised op already reported by V-5
    low, high = bounds
    if len(operands) < low or (high is not None and len(operands) > high):
        report.fail("V-31", f"{code}: SetOp {node.get('op')!r} has "
                            f"{len(operands)} operand(s)")


def check_substr(code, node, report: Report) -> None:
    """V-31 — `start` and `length` are plain integers or null/absent (§4.15)."""
    for member in ("start", "length"):
        if member in node and node[member] is not None:
            value = node[member]
            if not isinstance(value, int) or isinstance(value, bool):
                report.fail("V-31", f"{code}: SubstrOp {member} is {value!r}, "
                                    f"expected a JSON integer or null")


def check_rename(code, node, report: Report) -> None:
    """V-31 — rename pairs are simultaneous and disjoint (§4.24)."""
    clauses = node.get("clauses")
    if not isinstance(clauses, list) or not clauses:
        report.fail("V-31", f"{code}: RenameClauseOp with empty `clauses`")
        return
    sources, targets = set(), set()
    for clause in clauses:
        if not isinstance(clause, dict) or set(clause) != {"from_component",
                                                           "to_component"}:
            report.fail("V-31", f"{code}: rename clause {clause!r} is malformed")
            continue
        sources.add(clause["from_component"])
        targets.add(clause["to_component"])
    overlap = sources & targets
    if overlap:
        report.fail("V-31", f"{code}: rename sources and targets overlap on "
                            f"{sorted(overlap)}")


def _matches_declared_type(value, declared: str) -> bool:
    """§4.4, §4.8.1 — does `value` have the JSON type `declared` requires?

    `declared` is a Scalar param_type name ("Number", …) or "Item"; used both
    for a plain ParameterRef.default and for each element of a Set default.
    """
    if declared == "Item":
        return isinstance(value, str)
    expected = PARAMETER_DEFAULT_TYPES[declared]
    ok = isinstance(value, expected)
    if declared != "Boolean" and isinstance(value, bool):
        ok = False                      # bool is a subclass of int in Python
    if declared == "Number" and isinstance(value, int) \
            and not isinstance(value, bool):
        ok = True                       # a JSON integer is a valid Number
    if declared == "Date" and isinstance(value, str):
        ok = bool(ISO_DATE.match(value))
    return ok


def check_parameter_ref(code, node, parameters, report: Report) -> None:
    """V-29 — a ParameterRef agrees with its declaration (§3.9, §4.8.1)."""
    name = node.get("code")
    declared = node.get("param_type")
    if name not in parameters:
        report.fail("V-29", f"{code}: ParameterRef {name!r} is not declared "
                            f"in `parameters`")
    elif parameters[name] != declared:
        report.fail("V-29", f"{code}: ParameterRef {name!r} has param_type "
                            f"{declared!r}, declared {parameters[name]!r}")

    # §4.8.1 — `default` is REQUIRED and nullable, and a non-null value must
    # suit the declared type. Item and Set defaults are bare values (or
    # arrays of them), never a Scalar/Set Node.
    if "default" not in node:
        return                          # V-5 reports the missing member
    fallback = node["default"]
    if fallback is None:
        return
    if declared == "Item":
        if not _matches_declared_type(fallback, "Item"):
            report.fail("V-29", f"{code}: ParameterRef {name!r} of type "
                                f"\"Item\" has a default that is not a String")
        return
    if declared.startswith("Set"):
        # §4.8.1 — a set default is a JSON array of bare values of the
        # element type.
        if not isinstance(fallback, list):
            report.fail("V-29", f"{code}: ParameterRef {name!r} of type "
                                f"{declared!r} has a default that is not an array")
            return
        element_type = "Item" if declared == "SetItem" else declared[len("Set"):]
        bad = [v for v in fallback if not _matches_declared_type(v, element_type)]
        if bad:
            report.fail("V-29", f"{code}: ParameterRef {name!r} default "
                                f"array has element(s) {bad!r} not of type "
                                f"{element_type!r}")
        return
    if declared not in PARAMETER_DEFAULT_TYPES:
        return                          # unknown type already reported
    if not _matches_declared_type(fallback, declared):
        report.fail("V-29", f"{code}: ParameterRef {name!r} of type "
                            f"{declared!r} has default {fallback!r}")


def check_dimension_placement(code, root, report: Report) -> None:
    """V-22 — the four positions a Dimension may occupy (§4.7)."""
    condition_holders = ("WhereClauseOp", "SubClauseOp", "FilterOp")

    def descend(node, permitted: bool) -> None:
        cls = node["class_name"]
        if cls == "Dimension" and not permitted:
            report.fail("V-22", f"{code}: Dimension outside a clause "
                                f"condition or a SubstrOp operand")
        for member, child in child_nodes(node):
            enters = ((cls in condition_holders and member == "condition")
                      or (cls == "SubstrOp" and member == "operand"))
            descend(child, permitted or enters)
    descend(root, False)


#: §3.7.1 — Node classes and (class, op) pairs whose result is Boolean.
def check_preconditions(module, operations, tables, foreign_tables,
                        local_types, foreign_types, open_keys, parameters,
                        cell_of_ref, report: Report) -> None:
    declared = module["precondition_variables"]
    all_types = {**local_types, **foreign_types}
    for code, precondition in sorted(module["preconditions"].items()):
        # V-3
        if precondition.get("code") != code:
            report.fail("V-3", f"preconditions/{code}: code member is "
                               f"{precondition.get('code')!r}")
        if set(precondition) != PRECONDITION_MEMBERS:
            missing = PRECONDITION_MEMBERS - set(precondition)
            extra = set(precondition) - PRECONDITION_MEMBERS
            if missing:
                report.fail("V-3", f"preconditions/{code}: missing "
                                   f"{sorted(missing)}")
            if extra:
                report.fail("V-3", f"preconditions/{code}: unrecognised "
                                   f"{sorted(extra)}")
            if "ast" in missing:
                continue

        root = precondition["ast"]
        label = f"preconditions/{code}"

        # §3.7.1 — a precondition tree is an Operation tree, so every per-Node
        # rule applies. No Node class is excluded from, or reserved to, one.
        check_tree(label, root, tables, foreign_tables, local_types,
                   foreign_types, open_keys, parameters, declared,
                   cell_of_ref, report)

        # V-12 — the tree must not vary on any axis: its index set is either
        # empty or the single all-absent index. It is the axes that decide,
        # not the cardinality. That the result is Boolean is a semantic
        # constraint on the expression, settled before a document is produced
        # (Information Model §5.1), so it is deliberately not checked here.
        positions = index_set(root, report, label)
        if positions is not OPAQUE and axes_of(positions):
            report.fail("V-12", f"{label}: tree varies on axes "
                                f"{sorted(axes_of(positions))} and yields "
                                f"{len(positions)} results; a precondition must "
                                f"yield a Scalar")
        # V-13
        affected = precondition.get("affected_operations") or []
        if not affected:
            report.fail("V-13", f"{code}: empty affected_operations")
        for target in affected:
            if target not in operations:
                report.fail("V-13", f"{code}: affected operation {target!r} "
                                    f"is not in `operations`")


def check_dependencies(module, uri, operations, dep_modules,
                       report: Report) -> None:
    info = module["dependency_information"]

    # V-27 — all three members are REQUIRED (§3.11); an empty
    # `alternative_dependencies` is emitted rather than omitted.
    DEPENDENCY_MEMBERS = {"intra_instance_validations",
                          "cross_instance_dependencies",
                          "alternative_dependencies"}
    missing = DEPENDENCY_MEMBERS - set(info)
    extra = set(info) - DEPENDENCY_MEMBERS
    if missing:
        report.fail("V-27", f"dependency_information missing {sorted(missing)}")
    if extra:
        report.fail("V-27", f"dependency_information has unrecognised "
                            f"member(s) {sorted(extra)}")
    if {"intra_instance_validations", "cross_instance_dependencies"} & missing:
        return

    # V-28
    seen_in_group: dict[str, int] = {}
    for position, group in enumerate(info.get("alternative_dependencies", [])):
        if len(group) < 2:
            report.fail("V-28", f"alternative_dependencies[{position}] has "
                                f"{len(group)} member(s), expected at least 2")
        for target in group:
            if target not in dep_modules:
                report.fail("V-28", f"alternative_dependencies[{position}]: "
                                    f"{target!r} is not a `dependency_modules` member")
            if target in seen_in_group and seen_in_group[target] != position:
                report.fail("V-28", f"{target!r} appears in groups "
                                    f"{seen_in_group[target]} and {position}")
            seen_in_group[target] = position

    intra = set(info["intra_instance_validations"])
    cross: set[str] = set()
    referenced: set[str] = set()

    for position, entry in enumerate(info["cross_instance_dependencies"]):
        label = f"cross_instance_dependencies[{position}]"
        # V-13
        affected = entry.get("affected_operations") or []
        if not affected:
            report.fail("V-13", f"{label}: empty affected_operations")
        for target in affected:
            if target not in operations:
                report.fail("V-13", f"{label}: affected operation {target!r} "
                                    f"is not in `operations`")
            cross.add(target)
        # V-16
        modules = entry.get("modules") or []
        if not modules:
            report.fail("V-16", f"{label}: empty modules")
        for reference in modules:
            target_uri = reference.get("URI")
            referenced.add(target_uri)
            if target_uri == uri:
                if reference.get("ref_period") == "T":
                    report.fail("V-16", f"{label}: self-dependency with "
                                        f"ref_period 'T'")
            elif target_uri not in dep_modules:
                report.fail("V-16", f"{label}: URI {target_uri!r} is neither "
                                    f"a `dependency_modules` member nor the "
                                    f"root URI")
            if not REF_PERIOD.match(reference.get("ref_period") or ""):
                report.fail("V-16", f"{label}: ref_period "
                                    f"{reference.get('ref_period')!r}")
        # V-18 — an open-ended upper bound is the empty string here (§3.4).
        start, end = entry.get("from_reference_date"), entry.get("to_reference_date")
        if not ISO_DATE.match(start or ""):
            report.fail("V-18", f"{label}: from_reference_date {start!r}")
        if end not in (None, ""):
            if not ISO_DATE.match(end):
                report.fail("V-18", f"{label}: to_reference_date {end!r}")
            elif start and end < start:
                report.fail("V-18", f"{label}: to_reference_date {end} < "
                                    f"from_reference_date {start}")

    # V-14
    both = intra & cross
    if both:
        report.fail("V-14", f"{len(both)} Operation(s) in both "
                            f"`intra_instance_validations` and "
                            f"`cross_instance_dependencies`, e.g. "
                            f"{sorted(both)[0]}")
    uncovered = set(operations) - intra - cross
    if uncovered:
        report.fail("V-14", f"{len(uncovered)} Operation(s) in neither group, "
                            f"e.g. {sorted(uncovered)[0]}")
    unknown = (intra | cross) - set(operations)
    if unknown:
        report.fail("V-14", f"{len(unknown)} listed code(s) not in "
                            f"`operations`, e.g. {sorted(unknown)[0]}")
    # V-16
    unused = set(dep_modules) - referenced
    if unused:
        report.fail("V-16", f"{len(unused)} `dependency_modules` member(s) "
                            f"never referenced, e.g. {sorted(unused)[0]}")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def validate(path: Path) -> Report:
    report = Report()
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        report.fail("V-1", f"root is a JSON {type(raw).__name__}, "
                           f"expected an object")
        return report
    check_document(raw, report)
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Validate DPM-XL Script Documents against Script Format "
                    "§5.3 (rules V-1 .. V-32).")
    parser.add_argument("files", nargs="+", type=Path,
                        help="Script Document(s) to validate")
    parser.add_argument("--json", action="store_true",
                        help="emit machine-readable results")
    parser.add_argument("--quiet", action="store_true",
                        help="suppress output; report through the exit code")
    parser.add_argument("--max-examples", type=int, default=3, metavar="N",
                        help="violations to show per rule (default 3)")
    args = parser.parse_args(argv)

    results = {}
    failed = False
    for path in args.files:
        try:
            results[str(path)] = validate(path)
        except (OSError, json.JSONDecodeError) as error:
            if not args.quiet:
                print(f"{path}: cannot read: {error}", file=sys.stderr)
            return 2
        if not results[str(path)].valid:
            failed = True

    if args.quiet:
        return 1 if failed else 0

    if args.json:
        print(json.dumps({
            name: {"valid": report.valid,
                   "violations": {rule: report.violations[rule]
                                  for rule in report.rules()},
                   "unverified_operators": report.unverified}
            for name, report in results.items()
        }, indent=2))
        return 1 if failed else 0

    width = max(len(name) for name in results)
    for name, report in results.items():
        status = "VALID" if report.valid else f"{report.count} violation(s)"
        print(f"{name:<{width}}  {status}")
        for rule in report.rules():
            details = report.violations[rule]
            print(f"    {rule}  ({len(details)})")
            for detail in details[:args.max_examples]:
                print(f"        {detail}")
            if len(details) > args.max_examples:
                print(f"        ... and {len(details) - args.max_examples} more")

    total = sum(report.count for report in results.values())
    unverified = sum(report.unverified for report in results.values())
    print()
    print(f"{len(results)} document(s), {total} violation(s)"
          + (f", {unverified} Operator(s) not statically verifiable"
             if unverified else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

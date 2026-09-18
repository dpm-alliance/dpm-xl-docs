# 3 DPM-XL Data Selection

The data for the expression language is selected by referencing the DPM *Cells* and *Key Variables* of DPM Report Tables.

The references to the *Cells* and *Key Variables* are done by means of the table codes, ordinate codes and the codes of the Properties associated to the *Key Variables*.

The data selection is done by the interaction of the selection *Operator*, represented by the curly brackets symbol, and the general scope of an expression, provided by the with *Operator*. With this interaction, it is possible to individuate individual DPM *Variables* or sets of them. The actual data are referring to DPM *Variables*, thus making it possible to obtain the data required for the calculations.

The selection of data, together with the DPM definitions, make it possible to determine the *Structure* of the *Recordset* on compile-time.

## 3.1 Parameter Selection

A *Parameter* is an external input to an *Operation*: a value of a declared *Data Type* (or a set of such values) that is provided as input to the runtime alongside the data instances. Using a *Parameter* in an *Operation* therefore implies that the execution environment must supply a value for the *Parameter* together with the data instances; otherwise, the *Parameter* takes its declared default (or `null` when no default is declared).

*Parameters* let the same *Operation* be reused across executions that differ only in a threshold, a reference date, an enumeration member, an allowed set of codes, or any other configuration value that is not stored in the *Cells* of the input tables.

*Parameters* are referenced using the selection *Operator* with the `p` prefix, followed by the declared *Data Type*, and optionally followed by a default value.

### 3.1.1 Syntax

> { **p**[_]parameter, **type** [, _default_value_] }

where:

- **parameter** is the parameter code, following the same conventions as a *Variable* code. As with the other selection prefixes, an optional underscore between `p` and the code is allowed as a cosmetic separator: `pthreshold` and `p_threshold` refer to the same *Parameter*. To reference a code that itself starts with an underscore, backtick-escape the code (e.g. `` p`_legacy_param` ``).
- **type** is the declared *Data Type* of the *Parameter*. Allowed values:
    - *Scalar* types: `number`, `integer`, `string`, `date`, `boolean`, `item`.
    - *Set* types: `set-number`, `set-integer`, `set-string`, `set-date`, `set-boolean`, `set-item` — a *Set* parameter takes a set of values of the corresponding *Scalar* type at execution time.
- **default_value** is an optional default, written as `default: <literal>` for *Scalar* types, `default: [namespace:code]` for parameters of type `item`, or `default: { <literal>, ... }` / `default: { [ns:c], ... }` for *Set* types. When omitted, the implicit default is `null`.

### 3.1.2 Semantics

A *Parameter Selection* evaluates to a *Scalar* (for the basic types) or a set value (for the `set-*` types), of the declared *Data Type*. The value is supplied by the execution environment at run time, as part of the input alongside the data instances. If no value is supplied:

- if a `default` is declared in the selection, the *Parameter* takes the default value;
- otherwise, the *Parameter* is `null`.

The declared type fixes the *Data Type* of the resulting value and is used at compile time to check that the *Parameter* is used in a position compatible with its type (e.g. an `integer` *Parameter* cannot be passed where a `date` is required, unless an explicit cast is applied; a `set-item` *Parameter* can be used as the right-hand side of an `in` *Operator*, but not where a *Scalar* is expected).

### 3.1.3 Additional Constraints

- The default value, when provided, must be compatible with the declared *Data Type*: an `integer` *Parameter* accepts an integer literal default; a `date` *Parameter* accepts a date literal default; an `item` *Parameter* accepts an item reference default; a `set-<type>` *Parameter* accepts a set literal whose elements are of `<type>`.
- The declared *Data Type* must be one of the twelve listed above. *Number Interval*, *Subcategory* and *Recordset*-shaped *Parameters* are not supported by this syntax.

### 3.1.4 Runtime Binding

Because a *Parameter* is not part of the data instances, **using a *Parameter* in an *Operation* changes the contract between the *Operation* and its execution environment**: the engine must accept the bindings for every *Parameter* referenced by the *Operation* (and by any other *Operation* it transitively depends on), in addition to the usual data instances. How those bindings are delivered (configuration file, command-line argument, REST parameter, etc.) is left to the implementer; this specification only requires that the engine resolve each *Parameter* code to a value of the declared type at evaluation time, falling back to the declared default — or to `null` — when no binding is supplied.

### 3.1.5 Examples

Declaring a *Parameter* with no default — evaluates to `null` if no binding is supplied at runtime:

```
{p_threshold, number}
```

Declaring a *Parameter* with an explicit default:

```
{p_threshold, number, default: 0}
{p_cutoff, date, default: #2025-01-01#}
{p_ccy, item, default: [eba_CU:EUR]}
```

Using a *Parameter* in an *Operation* — the same expression can be re-executed against different thresholds without modifying the DPM:

```
{tT1, r010, c010} > {p_threshold, number, default: 0}
```

Using a *Parameter* inside a `where` clause to filter *Records* by an externally-supplied reference date:

```
{tT1}[where refPeriod >= {p_cutoff, date}]
```

Using an Item-typed *Parameter* with a default — filters *Records* by currency, defaulting to `EUR` when no binding is supplied:

```
{tT1}[where ccy = {p_ccy, item, default: [eba_CU:EUR]}]
```

Declaring a *Set* *Parameter* and using it as the right-hand side of `in` — the runtime supplies the set of allowed currencies, falling back to a single-element default when unbound:

```
{tT1}[where ccy in {p_ccys, set-item, default: {[eba_CU:EUR]}}]
```

Referencing a parameter whose code starts with an underscore via backtick escaping:

```
{p`_legacy_param`, number}
```
# 7 Aggregate Operators

## 7.1 Aggregate Operators' General Behaviour

Aggregate *Operators* apply a summarising computation (sum, count, average, …) to the *Fact* values of the *Operand* *Recordset*. They support two mutually exclusive invocation styles:

- A **reducing invocation** — the `group by` clause, or no clause at all — which collapses each group of *Records* into a single result *Record*: N *Records* per group become one. This is the classic aggregation and is the subject of §7.1.1–§7.1.7.
- An **analytical invocation** — the `over` clause — which computes the same aggregate over a *window* of related *Records* but returns **one output *Record* per input *Record***, so the *Recordset* keeps its cardinality. Analytical invocations are the unifying mechanism behind moving averages, running totals, shares of a parent total and ranks. They are described in [§7.1.8](#718-the-analytical-invocation-over-clause).

The `group by` and `over` clauses are mutually exclusive: an invocation uses either `group by` (or nothing) or `over`, never both.

### 7.1.1 Syntax

> **aggregateOperator (** op {**group by** groupingId {, groupingId}\*} **)**

> **aggregateOperator (** op **over (** analyticClause **) )**

### 7.1.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.1.3 Output

> rset | scal <*>

### 7.1.4 Semantics
In the **reducing invocation**, aggregate *Operators* perform operations on the measures of the *Operand* *Recordset*, calculating the required aggregated values for groups of *Records*. The groups of *Records* to be aggregated are specified through the grouping clause. If no grouping clause is used, the operation shall be calculated on all the *Records*, resulting in a *Scalar*.

In the **analytical invocation**, the same computation is applied without reducing: for each input *Record* the *Operator* is evaluated over the *window* of *Records* defined by the `over` clause, and the result replaces that *Record*'s *Fact* value. The output has the same *Key Components* and the same number of *Records* as the input (see [§7.1.8](#718-the-analytical-invocation-over-clause)).

### 7.1.5 Additional Constraints
The allowed *Data Types* depend on the specific *Operator* according to the following table:

| *Operator*      | *Operand* Type      | Result Type      |
|---------------|-------------------|------------------|
| Sum           | Number            | Number           |
| Count         | Any               | Integer          |
| Min           | Any               | Any              |
| Max           | Any               | Any              |
| Average       | Number            | Number           |
| Median        | Number            | Number           |
| Rank          | Any               | Integer          |

For **Average** and **Median**, when all input values are _Integer_, the result is implicitly promoted to _Number_ — the result type is always _Number_, even on integer-only inputs (e.g., `avg(1, 2)` returns `1.5`; `median(1, 2, 3, 4)` returns `2.5`). For **Sum**, an _Integer_-only input yields an _Integer_ result. **Count** always returns _Integer_ by definition. **Min** and **Max** return one of the input values, so an _Integer_-only input yields an _Integer_ result. **Rank** ([§7.8](#78-rank-rank)) exists only in the analytical invocation and always returns _Integer_.

The *Components* named in the grouping clause, and in the `partition by` and `order by` sub-clauses of the `over` clause, shall be present in the *Operand*.

### 7.1.6 Behaviour
In the **reducing invocation**, aggregate *Operations* generate a *Recordset* or a *Scalar*, depending on the grouping clause.

If the grouping clause exists, the structure of the resulting *Recordset* has as *Key Components* the *Components* in the group by

The result may be:

- A *Recordset*, for which the resulting structure contains as *Key* dimensions the *Components* included in the group by clause.
- A *Scalar*, if the grouping clause is omitted.

In the **analytical invocation** (`over` clause), the result is always a *Recordset* with the same *Key Components* and the same number of *Records* as the *Operand* (see [§7.1.8](#718-the-analytical-invocation-over-clause)).

Null values are excluded from the calculations. If all input values (in a group, or in a *Record*'s window) are null, the result is null. The only exception is **Count**, which returns 0 in that case (see below).

#### 7.1.6.1 Empty Operands

An aggregation has no values to work with when its *Operand* contains no *Records*, or when every one of its input values has been excluded as null. In that case:

- **Count** returns 0. Counting is a question about the collection rather than about the values, and the empty collection has a well-defined answer. **Count** therefore never returns null: its result is always a non-negative Integer.
- **Sum**, **Min**, **Max**, **Average** and **Median** return null. These *Operators* are defined over the input values, and with no values there is no result to report.

In particular, **Sum** over an empty *Operand* is null and not 0. Although 0 is the identity element of addition, returning it would state that a total of zero was reported, when in fact nothing was reported at all. This also keeps the empty case consistent with the all-null case: since null values are excluded before the calculation, an all-null *Operand* is an empty *Operand*, and both yield null.

When a grouping clause is present, the exclusion of null values takes place within each group and not before the grouping. The groups are derived from the *Records* present in the *Operand*, and a *Record* whose value is null still determines a group. A group all of whose values are null therefore does exist, and the rules above apply to it: **Count** returns 0 for that group and the other *Operators* return null for it. Retaining the group preserves the information that the *Records* were there, so that the nvl *Operator* or any subsequent *Operation* can act on the result.

What cannot happen is a group emerging for a *Key* combination that is not present in the *Operand*. The values of the *Key Components* of a group have to be taken from a *Record*, and *Key Components* cannot be empty, so no *Record* can be produced for a combination that was never reported: DPM-XL does not build the Cartesian product of all the potential combinations of *Key* values (see [Implicit Nulls Generation](../information-model/04-null-treatment.md#421-implicit-nulls-generation), where implicit nulls are generated only for the *Key* combinations having at least one value reported). Accordingly, an *Operand* with no *Records* at all yields an empty *Recordset*, and not a *Recordset* holding a *Record* with the value 0 or null.

Whether a null is explicit or implicit makes no difference to any of the above, in line with [Implicit and Explicit Null Values](../information-model/04-null-treatment.md#42-implicit-and-explicit-null-values). What determines the result is whether a *Record* exists in the *Operand*, not the reason why its value is null.

### 7.1.7 Examples
Supposing the following *Recordset* with name rs1:

| r {#k} | c {#k} | CNT {#k} | f {#f} |
|--------|--------|----------|--------|
| 010    | 010    | PT       | 100    |
| 010    | 020    | PT       | 200    |
| 020    | 010    | PT       | 300    |
| 020    | 020    | PT       | 400    |
| 010    | 010    | DE       | 500    |
| 010    | 020    | DE       | 600    |
| 020    | 010    | DE       | 700    |
| 020    | 020    | DE       | 800    |

`sum(rs1 group by r)` results in:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 1400   |
| 020    | 2200   |

`sum(rs1 group by r, CNT)` results in:

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 010    | PT       | 300    |
| 020    | PT       | 700    |
| 010    | DE       | 1100   |
| 020    | DE       | 1500   |

`count(rs1)` results in: `8` (*Scalar*)

Supposing now a *Recordset* with name rs2 whose selection yields no *Records* at all:

- `count(rs2)` results in: `0` (*Scalar*)
- `sum(rs2)` results in: null (*Scalar*)
- `sum(rs2 group by r)` results in an empty *Recordset* with *Key Component* r and no *Records*.
- `count(rs2 group by r)` results in the same empty *Recordset*, and not in a *Record* holding the value 0. Under a grouping clause there is no *Record* to provide a value for r, so the 0 has nowhere to appear. **Count** and **Sum** thus give the same result here, and only differ in the *Scalar* case above.

Supposing instead a *Recordset* with name rs3 having the same *Records* as rs1, but with all the values of f null:

- `count(rs3)` results in: `0` (*Scalar*)
- `sum(rs3)` results in: null (*Scalar*)
- `sum(rs3 group by r)` results in:

| r {#k} | f {#f} |
|--------|--------|
| 010    | null   |
| 020    | null   |

- `count(rs3 group by r)` results in:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 0      |
| 020    | 0      |

The groups for r are still formed, because the *Records* are present in rs3 and provide the values of r. This is the only way a **Count** of 0 can appear in a grouped result.

### 7.1.8 The Analytical Invocation (over clause)

The `over` clause turns an aggregate *Operator* into an **analytical invocation** (a *window function* in SQL terms). Instead of collapsing groups, the *Operator* computes its result over a *window* of *Records* defined relative to each *Record*, and emits one output *Record* per input *Record*.

#### 7.1.8.1 Syntax

> op **over (**
> {**partition by** partitionId {, partitionId}\*}
> {**order by** orderId {**asc**|**desc**} {, orderId {**asc**|**desc**}}\*}
> {windowClause}
> **)**

where

> windowClause: **(** **data points** | **range** **)** **between** windowBoundary **and** windowBoundary

> windowBoundary: **unbounded preceding** | **unbounded following** | **current data point** | int {time_period_code} **preceding** | int {time_period_code} **following**

> partitionId, orderId: scal <prop\>

> int: A positive *Integer* literal.

> time_period_code: One of the [time-period codes](./10-time-operators.md#1012-input-parameters) `A`, `S`, `Q`, `M`, `W`, `D`.

All three inner sub-clauses are optional, subject to the constraints below. `asc` (ascending) is the default direction and may be omitted.

#### 7.1.8.2 Partitioning (partition by)

`partition by` splits the *Operand* *Recordset* into groups (*partitions*) that share the same values for the listed *Components*, analogously to the *Key Components* of a reducing `group by`. Every *Record*'s window is confined to its own partition. When `partition by` is omitted, the whole *Recordset* is a single partition.

Unlike `group by`, `partition by` does **not** reduce: the partition *Components* are not the *Key* of the result — the result keeps the *Key Components* of the *Operand*.

#### 7.1.8.3 Ordering (order by)

*Records* are an unordered collection (see [§2.2.3](../information-model/02-information-model.md#223-recordsets)), so any order-dependent window needs an explicit ordering. `order by` establishes, within each partition, the order in which *Records* are considered. Each ordering *Component* may be followed by `asc` (ascending, the default) or `desc` (descending).

`order by` is:

- **required** for order-dependent *Operators* (`rank`) and whenever a `data points` or `range` window clause is used;
- **optional** otherwise — when omitted the whole partition forms one unordered window and every *Record* in the partition receives the same aggregate.

These requirements are stated as constraints in [§7.1.8.6](#7186-additional-constraints).

#### 7.1.8.4 Window Clause (data points / range)

The window clause narrows the window to a band of *Records* around the current one, expressed as two bounds. A window clause requires an `order by`.

Each bound is a *windowBoundary*, written in the direction it travels from the current *Record*:

- **`current data point`** — the current *Record* itself.
- **`n preceding`** / **`n following`** — `n` units before or after the current *Record*, where `n` is a positive *Integer*. The zero offset is written `current data point`.
- **`unbounded preceding`** / **`unbounded following`** — the start or the end of the partition.

The unit of `n` depends on the keyword introducing the clause:

- **`data points`** counts **positional offsets**: `n` is a number of *Records* in the `order by` order, so `1 preceding` is the immediately preceding *Record*. Use `data points` when position — not the value of the order key — is what matters.
- **`range`** counts **value offsets** on a single `order by` *Component*: the window contains every *Record* whose order value lies within `n` of the current *Record*'s order value. Use `range` when the ordering has gaps, because the window is defined by the order value rather than by record position.

Where the `order by` *Component* is of *Date* type, `n` is followed by a [time-period code](./10-time-operators.md#1012-input-parameters) giving the unit of the offset, so that `11 M preceding` is the *Record* eleven months before the current one. The offset is computed exactly as in [time_shift](./10-time-operators.md#101-time-shift): `n` units of the given period are subtracted from (`preceding`) or added to (`following`) the current *Record*'s order value, and the bound is that resulting date. Where the `order by` *Component* is numeric, `n` is a plain number of units of that *Component* and no period code is written.

`preceding` and `following` are always relative to the **`order by` order**, not to the natural order of the values: a `preceding` bound lies toward the start of the ordering and a `following` bound toward its end. Under a `desc` ordering the start of the ordering holds the *greatest* order values, so a `preceding` bound then has a greater order value than the current *Record* — a later date, where the order key is a *Date*. This is the same rule `data points` follows, where `1 preceding` is the previous *Record* of the ordering whichever direction it runs.

> A **trailing** window therefore requires an **ascending** ordering. `order by refDate range between 11 M preceding and current data point` covers the twelve months ending at the current *Record*, whereas the same window clause under `order by refDate desc` reaches forward instead, over the twelve months beginning at it.

When the window clause is omitted, the window is the **whole partition** (still honouring `order by` where it matters, as for `rank`). A time-period code may be written only on a `range` bound whose `order by` *Component* is of *Date* type; it is not allowed on a `data points` bound, whose unit is always one *Record*. The bounds are further constrained as set out in [§7.1.8.6](#7186-additional-constraints).

#### 7.1.8.5 Structure and Records

The result is a *Recordset* with the same *Key Components* and the same number of *Records* as the *Operand*. Each *Record*'s *Fact* value is replaced by the aggregate computed over that *Record*'s window; *Key Components* are unchanged and *Attribute Components* are not propagated. Null values inside a window are excluded from the computation, consistent with the reducing invocation; if a window contains only nulls, the result for that *Record* follows the empty-*Operand* rules of [§7.1.6.1](#7161-empty-operands) — 0 for **Count**, null for the other *Operators*. The result *Data Type* of each *Operator* is the one given in [§7.1.5](#715-additional-constraints).

#### 7.1.8.6 Additional Constraints

The *Components* named in `partition by` and in `order by` shall be present in the *Operand*, as required by [§7.1.5](#715-additional-constraints).

**The offset `n`.** The `n` of `n preceding` or `n following` shall be at least 1. The zero offset — the current *Record* itself — is written `current data point`, which is its only form: `0 preceding`, `0 following` and `0 M preceding` are not allowed.

**When an `order by` is required.** `order by` is required in two cases: when a window clause is used, and always with the `rank`operator. Both narrow or read a sequence, and *Records* are an unordered collection ([§2.2.3](../information-model/02-information-model.md#223-recordsets)), so without an explicit ordering neither has a determined result. This holds even for `between unbounded preceding and unbounded following`, whose meaning does not in fact depend on the ordering: the whole partition is expressed by omitting the window clause, which is the form to use.

**`rank` takes no window clause.** `rank` ranks by the `order by` value, and a window clause is defined along that same `order by`, positionally for `data points` and by value for `range`. A frame can therefore never give `rank` a different comparison group to measure against. To compare a *Record* against a window of its neighbours, use an aggregate over that window and compare it with the *Record*'s own value, as in `max_aggr(rs1 over (partition by qInstrument order by refDate range between 11 M preceding and current data point))`.

**The `order by` of a `range` clause.** Where the window clause is a `range`, the `order by` shall name exactly one *Component*, and that *Component* shall be of *Date*, *Number* or *Integer* type. A `range` bound measures an offset against the order *value*, and a lexicographic ordering over several *Components* offers no single value to offset from, while a *String* or *Boolean* *Component* offers no notion of lying within `n` of a value. Neither restriction applies to `data points`, which needs only a sequence, nor to `rank`; both accept several ordering *Components*.

**Ordering of the window bounds.** The bounds of a window clause are ordered by the position they denote relative to the current *Record*:

> `unbounded preceding` < `n preceding` (decreasing in `n`) < `current data point` < `n following` (increasing in `n`) < `unbounded following`

- The first bound shall not be after the second in this order. Where the two bounds denote the same position, the window is that single position — as in `between current data point and current data point`.
- `unbounded following` shall not be the first bound, and `unbounded preceding` shall not be the second. Such a window is either empty or a single *Record* of the partition repeated for every *Record*, which is never the intent: the whole partition is written `between unbounded preceding and unbounded following`, and a single aggregate over the partition is a reducing invocation ([§7.1.6](#716-behaviour)).

**Commensurable period codes.** Where both bounds of a `range` clause carry a [time-period code](./10-time-operators.md#1012-input-parameters), the two codes shall be commensurable: either both drawn from `A`, `S`, `Q`, `M`, or both from `W`, `D`. Within each of these groups every code is a whole number of the same base unit (twelve, six, three and one *months*; seven and one *days*), so the bounds can always be ordered. Across the groups they cannot: `1 M preceding` is earlier than `30 D preceding` in some months and later in others, so a frame such as `range between 3 M preceding and 40 D preceding` has no determinate meaning and is not allowed.

#### 7.1.8.7 Examples

Considering the following *Recordset* with name rs1, holding a monthly balance per instrument:

| qInstrument {#k} | refDate {#k} | f {#f} |
|------------------|--------------|--------|
| A                | 2022-01-01   | 100    |
| A                | 2022-02-01   | 200    |
| A                | 2022-03-01   | 300    |
| B                | 2022-01-01   | 10     |
| B                | 2022-02-01   | 20     |
| B                | 2022-03-01   | 30     |

**Trailing 2-month moving average** — `avg(rs1 over (partition by qInstrument order by refDate data points between 1 preceding and current data point))` averages each *Record* with the immediately preceding one in its partition:

| qInstrument {#k} | refDate {#k} | f {#f} |
|------------------|--------------|--------|
| A                | 2022-01-01   | 100    |
| A                | 2022-02-01   | 150    |
| A                | 2022-03-01   | 250    |
| B                | 2022-01-01   | 10     |
| B                | 2022-02-01   | 15     |
| B                | 2022-03-01   | 25     |

(The first *Record* of each partition has no predecessor, so its window is just itself.)

**Running total** — `sum(rs1 over (partition by qInstrument order by refDate data points between unbounded preceding and current data point))` accumulates from the start of each partition up to the current *Record*:

| qInstrument {#k} | refDate {#k} | f {#f} |
|------------------|--------------|--------|
| A                | 2022-01-01   | 100    |
| A                | 2022-02-01   | 300    |
| A                | 2022-03-01   | 600    |
| B                | 2022-01-01   | 10     |
| B                | 2022-02-01   | 30     |
| B                | 2022-03-01   | 60     |

**12-month moving average, gap-tolerant** — `avg(rs1 over (partition by qInstrument order by refDate range between 11 M preceding and current data point))` averages every *Record* whose `refDate` falls within the 12 calendar months ending at the current *Record*. Because the bound is value-based, a missing month simply contributes no *Record* to the window; the positional form `data points between 11 preceding and current data point` would instead always reach 11 *Records* back, ignoring calendar gaps.

## 7.2 Sum (sum)

### 7.2.1 Syntax

> **sum(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.2.2 Input Parameters

> op: rset <num + interval\>

> groupingId: scal <prop\>

### 7.2.3 Output

> rset | scal <num + interval\>

### 7.2.4 Semantics
Returns the sum of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:

- The centre is calculated as the sum of the all the centers of the *Operands*.
- The radius is calculated as the sum of all the radiuses of the *Operands*.

### 7.2.5 Additional Constraints
None.

### 7.2.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.3 Count (count)

### 7.3.1 Syntax

> **count(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.3.2 Input Parameters

> op: rset <*> | scalar_set <*>

>  groupingId: scal <prop\>

### 7.3.3 Output

> rset | scal <int\>

### 7.3.4 Semantics
Returns the number of non-null values in the *Recordset* or in each group of *Records*. Follows the general semantics of aggregate *Operators*: null values are excluded from the calculation, so *Records* whose value is null do not contribute to the count.

Unlike the other aggregate *Operators*, **count** never returns null. Its result is always a non-negative Integer: an *Operand* with no *Records*, and an *Operand* whose values are all null, both yield 0 (see [Empty *Operands*](#7161-empty-operands)).

When op is a `scalar_set`, `count` returns the **cardinality** of the set — the number of distinct elements it contains. As sets are duplicate-free and never contain `null` (see [Set *Operators*' general behaviour](./13-set-operators.md#131-set-operators-general-behaviour)), the count is taken over the distinct, non-null values. The `group by` clause is inapplicable when op is a `scalar_set`.

### 7.3.5 Additional Constraints
None.

### 7.3.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

Note: Aggregate *Operators* generally ignore null values. This behavior can be overridden by using the nvl *Operator*. Applied to **count**, `count(nvl(op, 0))` counts every *Record* of op, including those whose value is null.

### 7.3.7 Examples

- `count({1, 2, 3, 3, 2})` results in `3` — the set literal de-duplicates first, leaving three distinct elements.
- `count(set_of({tT1, r010-020}[get CNT]))` results in `4` — the number of distinct `CNT` values (see [set_of](./13-set-operators.md#132-set-of-set_of)).
- `count(union({1, 2}, {2, 3}))` results in `3`.
- `count({})` results in `0` — the cardinality of the empty set.

## 7.4 Minimum Value (min_aggr)

### 7.4.1 Syntax

> **min_aggr(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.4.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.4.3 Output

> rset | scal <*\>

### 7.4.4 Semantics
Returns the minimum value of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:

- Centre is the minimum of the centres of all the *Operands*.
- Radius is the radius of the *Operand* with the minimum centre. If multiple *Operands* share the same minimum centre, the largest radius among them is taken.

### 7.4.5 Additional Constraints
None.

### 7.4.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.5 Maximum Value (max_aggr)

### 7.5.1 Syntax

> **max_aggr(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.5.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.5.3 Output

> rset | scal <*\>

### 7.5.4 Semantics
Returns the maximum value of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:

- Centre is the maximum of the centres of all the *Operands*.
- Radius is the radius of the *Operand* with the maximum centre. If multiple *Operands* share the same maximum centre, the largest radius among them is taken.

### 7.5.5 Additional Constraints
None.

### 7.5.6 Behaviour
[Aggregate *Operators'* general behaviour.](#71-aggregate-operators-general-behaviour)

## 7.6 Average (avg)

### 7.6.1 Syntax

> **avg(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.6.2 Input Parameters

> op: rset <num + interval\>

> groupingId: scal <prop\>

### 7.6.3 Output

> rset | scal <num + interval\>

### 7.6.4 Semantics
Returns the average of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:

- The centre is calculated as the average of the centres of all the *Operands*.
- The radius is calculated as the average of the radii of all the *Operands*.

### 7.6.5 Additional Constraints
None.

### 7.6.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.7 Median Value (median)

### 7.7.1 Syntax

> **median(** op {**group by** groupingId {, groupingId}\* | **over (** analyticClause **)** } **)**

### 7.7.2 Input Parameters

> op: rset <num + interval\>

> groupingId: scal <prop\>

### 7.7.3 Output

> rset | scal <num + interval\>

### 7.7.4 Semantics
Returns the median of the input values. Follows the general semantics of aggregate *Operators*.

For intervals, the centre is used as the ordering criterion. The result is defined as follows:

- **Centre**: the median of the centres of all the *Operands* (using standard median arithmetic on the centre values).
- **Radius**:
    - If the median position corresponds to a single *Operand*, its radius is used.
    - If multiple *Operands* share the same median centre (tie at the median position), the largest radius among them is taken.
    - If the median is the arithmetic average of two *Operands* with distinct centres (even number of *Operands*), the radius is the arithmetic average of their radii.

### 7.7.5 Additional Constraints
None.

### 7.7.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.8 Rank (rank)

`rank` is an **analytical-only** *Operator*: it has no reducing counterpart and is defined solely through the `over` clause ([§7.1.8](#718-the-analytical-invocation-over-clause)).

### 7.8.1 Syntax

> **rank(** op **over (** analyticClause **) )**

### 7.8.2 Input Parameters

> op: rset <*>

> partitionId, orderId: scal <prop\>

### 7.8.3 Output

> rset <int\>

### 7.8.4 Semantics

Returns, for each *Record*, the 1-based rank of its `order by` value within its partition, following the `order by` direction (`asc`/`desc`). *Records* that tie on the order value receive the same rank, and the next distinct value's rank is increased by the number of tied *Records* (standard competition ranking: 1, 2, 2, 4, …).

`rank` is order-dependent, so the `over` clause **must** contain an `order by`, and it takes no window clause: `rank` always considers the whole partition. The reason is that the `order by` is both the ranking key and the axis along which a window clause would be defined, so a frame could only report the *Record*'s position inside the frame rather than its standing within a comparison group (see [§7.1.8.6](#7186-additional-constraints)).

*Records* whose order value is null sort last and share the final rank group.

### 7.8.5 Additional Constraints

- The `over` clause must contain an `order by`.
- `partition by` is optional; without it the whole *Recordset* is ranked as a single partition.
- A window clause (`data points` / `range`) is not allowed with `rank`.

### 7.8.6 Behaviour

[Aggregate *Operators'* analytical invocation behaviour](#718-the-analytical-invocation-over-clause): one output *Record* per input *Record*, with the *Fact Component* replaced by the Integer rank. *Key Components* are unchanged and *Attribute Components* are not propagated.

### 7.8.7 Examples

Considering the following *Recordset* with name rs1, holding total assets per entity within each country:

| qCountry {#k} | id {#k} | f {#f} |
|---------------|---------|--------|
| PT            | X       | 300    |
| PT            | Y       | 100    |
| PT            | Z       | 300    |
| DE            | P       | 500    |
| DE            | Q       | 200    |

**Rank within country by total assets** (descending) — `rank(rs1 over (partition by qCountry order by f desc))` returns:

| qCountry {#k} | id {#k} | f {#f} |
|---------------|---------|--------|
| PT            | X       | 1      |
| PT            | Z       | 1      |
| PT            | Y       | 3      |
| DE            | P       | 1      |
| DE            | Q       | 2      |

(Within PT, the two *Records* valued `300` tie at rank `1`, so the `100` *Record* takes rank `3`. The *Fact* value is replaced by the rank.)
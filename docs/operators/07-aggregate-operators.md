# 7 Aggregate Operators

## 7.1 Aggregate Operators' General Behaviour

### 7.1.1 Syntax

> **aggregateOperator (** op {**group by** groupingId {, groupingId}\*} **)**

### 7.1.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.1.3 Output

> rset | scal <*>

### 7.1.4 Semantics
Aggregate *Operators* perform operations on the measures of the *Operand* *Recordset*, calculating the required aggregated values for groups of *Records*. The groups of *Records* to be aggregated are specified through the grouping clause. If no grouping clause is used, the operation shall be calculated on all the *Records*, resulting in a *Scalar*.

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

The *Components* in the grouping by clause shall be present in the *Operand*.

### 7.1.6 Behaviour
Aggregate *Operations* generate a *Recordset* or a *Scalar*, depending on the grouping clause.

If the grouping clause exists, the structure of the resulting *Recordset* has as *Key Components* the *Components* in the group by

The result may be:

- A *Recordset*, for which the resulting structure contains as *Key* dimensions the *Components* included in the group by clause.
- A *Scalar*, if the grouping clause is omitted.

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

## 7.2 Sum (sum)

### 7.2.1 Syntax

> **sum(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.2.2 Input Parameters

> op: rset <num + interval\>

> groupingId: scal <prop\>

### 7.2.3 Output

> rset | scal <num\>

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

> **count(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.3.2 Input Parameters

> op: rset <*>

>  groupingId: scal <prop\>

### 7.3.3 Output

> rset | scal <num\>

### 7.3.4 Semantics
Returns the number of *Records* in the *Recordset* or groups of *Records*. Follows the general semantics of aggregate *Operators*.

### 7.3.5 Additional Constraints
None.

### 7.3.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

Note: Aggregate *Operators* generally ignore null values. This behavior can be overridden by using the nvl *Operator*.

## 7.4 Minimum Value (min_aggr)

### 7.4.1 Syntax

> **min_aggr(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.4.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.4.3 Output

> rset | scal <num\>

### 7.4.4 Semantics
Returns the minimum value of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:

- The centre is calculated as the minimum value of the all the centers of the *Operands*.
- The radius is the radius of the *Operand* that has the minimum centre.

### 7.4.5 Additional Constraints
None.

### 7.4.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.5 Maximum Value (max_aggr)

### 7.5.1 Syntax

> **max_aggr(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.5.2 Input Parameters

> op: rset <*>

> groupingId: scal <prop\>

### 7.5.3 Output

> rset | scal <num\>

### 7.5.4 Semantics
Returns the maximum value of the input values. Follows the general semantics of aggregate *Operators*.

For intervals:
- The centre is calculated as the maximum value of the all the centers of the *Operands*.
- The radius is the radius of the *Operand* that has the maximum centre.

### 7.5.5 Additional Constraints
None.

### 7.5.6 Behaviour
[Aggregate *Operators'* general behaviour.](#71-aggregate-operators-general-behaviour)

## 7.6 Average (avg)

### 7.6.1 Syntax

> **avg(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.6.2 Input Parameters

> op: rset <num\>

> groupingId: scal <prop\>

### 7.6.3 Output

> rset | scal <num\>

### 7.6.4 Semantics
Returns the average of the input values. Follows the general semantics of aggregate *Operators*.

### 7.6.5 Additional Constraints
None.

### 7.6.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).

## 7.7 Median Value (median)

### 7.7.1 Syntax

> **median(** op {**group by** groupingId {, groupingId}\*} **)**

### 7.7.2 Input Parameters

> op: rset <num\>

> groupingId: scal <prop\>

### 7.7.3 Output

> rset | scal <num\>

### 7.7.4 Semantics
Returns the median of the input values. Follows the general semantics of aggregate *Operators*.

### 7.7.5 Additional Constraints
None.

### 7.7.6 Behaviour
[Aggregate *Operators'* general behaviour](#71-aggregate-operators-general-behaviour).
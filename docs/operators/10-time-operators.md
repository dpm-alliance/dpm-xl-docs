# 10 Time Operators

## 10.1 Time Shift

### 10.1.1 Syntax

> **time_shift(** op, period, numberPeriods, {var} **)**

### 10.1.2 Input Parameters

> op: rset <*> | scal <date\>

> period: scal <str\>

> numberPeriods: scal <int\>

> var: scal <prop\>

### 10.1.3 Output

> rset <*> | scal <date\>

### 10.1.4 Semantics
Changes the dates of the var *Component* of the *Recordset* op by adding (or subtracting) numberPeriods of the given period unit.

### 10.1.5 Additional Constraints

- If op is a *Recordset*, then

  - var must be provided,
  - must belong to the *Recordset* op, and
  - must be of *Date* type.

- If op is a *Scalar*, then var must be omitted.

The period must have one of the following values:

- _A_ for year
- _S_ for semester
- _Q_ for quarter
- _M_ for month
- _W_ for week
- _D_ for day

### 10.1.6 Behaviour
Returns a *Recordset* with the same data structure and number of *Records* as the input op. The dates for the *Component* var are modified by adding the numberPeriods (subtracting, if negative) of the given period unit.

### 10.1.7 Examples
Considering the following *Recordset*, generated from the selection {tT1, r010-020, c010-020}:

| refPeriod {#k} | r {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2022-01-01   | 010    | 010    | 100    |
| 2022-01-01   | 010    | 020    | 200    |
| 2022-01-01   | 010    | 010    | 300    |
| 2022-01-01   | 010    | 020    | 400    |
| 2022-04-01   | 020    | 010    | 500    |
| 2022-04-01   | 020    | 020    | 600    |
| 2022-04-01   | 020    | 010    | 700    |
| 2022-04-01   | 020    | 020    | 800    |

`time_shift({tT1, r010-020, c010-020}, Q, 1, refPeriod)` returns:

| refPeriod {#k} | r {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2022-04-01   | 010    | 010    | 100    |
| 2022-04-01   | 010    | 020    | 200    |
| 2022-04-01   | 010    | 010    | 300    |
| 2022-04-01   | 010    | 020    | 400    |
| 2022-07-01   | 020    | 010    | 500    |
| 2022-07-01   | 020    | 020    | 600    |
| 2022-07-01   | 020    | 010    | 700    |
| 2022-07-01   | 020    | 020    | 800    |

`time_shift({tT1, r010-020, c010-020}, Q, -1, refPeriod)` returns:

| refPeriod {#k} | r {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2021-10-01   | 010    | 010    | 100    |
| 2021-10-01   | 010    | 020    | 200    |
| 2021-10-01   | 010    | 010    | 300    |
| 2021-10-01   | 010    | 020    | 400    |
| 2022-01-01   | 020    | 010    | 500    |
| 2022-01-01   | 020    | 020    | 600    |
| 2022-01-01   | 020    | 010    | 700    |
| 2022-01-01   | 020    | 020    | 800    |
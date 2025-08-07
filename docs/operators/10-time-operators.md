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

> rset <*>

### 10.1.4 Semantics
Changes the dates of the var *Component* of the *Recordset* op by adding (or subtracting) the numberPeriods of the period type.

### 10.1.5 Additional Constraints
The *Component* var must belong to the *Recordset* op, and has to be of Time interval type.

The period must have one of the following values:

- _A_ for year
- _S_ for semester
- _Q_ for quarter
- _M_ for month
- _W_ for week
- _D_ for day

### 10.1.6 Behaviour
Returns a *Recordset* with the same data structure and number of *Records* as the input op. The dates for the *Component* var are modified by adding the numberPeriods (subtracting, if negative) of the period type.

### 10.1.7 Examples
Considering the following *Recordset*, generated from the selection {tT1, r010-020, c010-020}:

| RefDate {#k} | r {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2022Q1       | 010    | 010    | 100    |
| 2022Q1       | 010    | 020    | 200    |
| 2022Q1       | 010    | 010    | 300    |
| 2022Q1       | 010    | 020    | 400    |
| 2022Q2       | 020    | 010    | 500    |
| 2022Q2       | 020    | 020    | 600    |
| 2022Q2       | 020    | 010    | 700    |
| 2022Q2       | 020    | 020    | 800    |

`time_shift({tT1, r010-020, c010-020}, RefDate, Q, 1)` returns:

| RefDate {#k} | R {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2022Q2       | 010    | 010    | 100    |
| 2022Q2       | 010    | 020    | 200    |
| 2022Q2       | 010    | 010    | 300    |
| 2022Q2       | 010    | 020    | 400    |
| 2022Q3       | 020    | 010    | 500    |
| 2022Q3       | 020    | 020    | 600    |
| 2022Q3       | 020    | 010    | 700    |
| 2022Q3       | 020    | 020    | 800    |

`time_shift({tT1, r010-020, c010-020}, RefDate, Q, -1)` returns:

| RefDate {#k} | r {#k} | c {#k} | f {#f} |
|--------------|--------|--------|--------|
| 2021Q4       | 010    | 010    | 100    |
| 2021Q4       | 010    | 020    | 200    |
| 2021Q4       | 010    | 010    | 300    |
| 2021Q4       | 010    | 020    | 400    |
| 2022Q1       | 020    | 010    | 500    |
| 2022Q1       | 020    | 020    | 600    |
| 2022Q1       | 020    | 010    | 700    |
| 2022Q1       | 020    | 020    | 800    |
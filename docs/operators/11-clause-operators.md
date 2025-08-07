# 11 Clause Operators

Clause *Operators* serve to perform operations on the DPM *Key Components* of *Recordsets*. Use of clause *Operators* with Standard *Key Components* are not allowed.

## 11.1 where

### 11.1.1 Syntax

> op **[where** condition **]**

### 11.1.2 Input Parameters

> op: rset <*>

> condition: expression <boo\>

### 11.1.3 Output

> rset <*>

### 11.1.4 Semantics
Filters a *Recordset* based on the value of a *Key Component*.

### 11.1.5 Additional Constraints
Condition must be a Boolean expression using as input *Key Components* of op.

### 11.1.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*, and with the *Records* resulting from the evaluation of the filtering condition. When the condition evaluates to true, the *Record* is kept, otherwise (including null), the *Record* is not kept.

### 11.1.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| RefDate {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q1       | 010    | DE       | 300    |
| 2022Q1       | 010    | IT       | 400    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |
| 2022Q2       | 020    | DE       | 700    |
| 2022Q2       | 020    |          | 800    |

`{tT1, r010-020}[where CNT in {"ES", "PT"}]` returns:

| RefDate {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |

## 11.2 rename

### 11.2.1 Syntax

> op **[rename** compFrom to compTo **]**

### 11.2.2 Input Parameters

> op: rset <*>

> compFrom: scal <prop\>

> compTo: scal <prop\>

### 11.2.3 Output

> rset <*>

### 11.2.4 Semantics
Changes the name of a *Component*.

### 11.2.5 Additional Constraints
compFrom must belong to op.

compFrom must be a DPM *Key Component* (i.e., cannot refer to row, column or sheet).

compTo name cannot be already used in op.

compTo name cannot be the name of a standard *Key Component* ("r", "c" or "s").

### 11.2.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*, except the renamed *Components*, which change name. All the *Records* are kept.

### 11.2.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| RefDate {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q1       | 010    | DE       | 300    |
| 2022Q1       | 010    | IT       | 400    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |
| 2022Q2       | 020    | DE       | 700    |
| 2022Q2       | 020    |          | 800    |

`{tT1, r010-020}[rename CNT to country]` returns:

| RefDate {#k} | r {#k} | country {#k} | f {#f} |
|--------------|--------|--------------|--------|
| 2022Q1       | 010    | ES           | 100    |
| 2022Q1       | 010    | PT           | 200    |
| 2022Q1       | 010    | DE           | 300    |
| 2022Q1       | 010    | IT           | 400    |
| 2022Q2       | 020    | ES           | 500    |
| 2022Q2       | 020    | PT           | 600    |
| 2022Q2       | 020    | DE           | 700    |
| 2022Q2       | 020    |              | 800    |

## 11.3 get

### 11.3.1 Syntax

> op **[get** component **]**

### 11.3.2 Input Parameters

> op: rset <*>

> component: scal <prop\>

### 11.3.3 Output

> rset <*>

### 11.3.4 Semantics
Returns a *Recordset* where the fact is the value of one of the *Key Components*.

### 11.3.5 Additional Constraints
component must belong to op.

### 11.3.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*. The *Fact* values the *Records* are substituted for the value of the selected *Component*.

*Records* with null *Fact* value will be omitted.

### 11.3.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| RefDate {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q1       | 010    | DE       | 300    |
| 2022Q1       | 010    | IT       | 400    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |
| 2022Q2       | 020    | DE       | 700    |
| 2022Q2       | 020    | IT       | 800    |

`{tT1, r010-020}[get CNT]` returns:

| RefDate {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | ES     |
| 2022Q1       | 010    | PT       | PT     |
| 2022Q1       | 010    | DE       | DE     |
| 2022Q1       | 010    | IT       | IT     |
| 2022Q2       | 020    | ES       | ES     |
| 2022Q2       | 020    | PT       | PT     |
| 2022Q2       | 020    | DE       | DE     |
| 2022Q2       | 020    | IT       | IT     |
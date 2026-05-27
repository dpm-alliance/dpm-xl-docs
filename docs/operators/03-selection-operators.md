# 3 Selection Operators

## 3.1 Selection Operator

### 3.1.1 Syntax

> {
  
>> [**t**table | **v**variable | **o**operation]1

>> [**r**row [, **r**row]* | **r**row_init–row_end | **r*** ]*

>> [**c**col [, **c**col]* | **c**col_init–col_end | **c*** ]*

>> [**s**sheet [, **s**sheet]* | **s**sheet_init–sheet_end | **s*** ]*

>> _interval_arithmetics_, _fallback_value_

>}

### 3.1.2 Input Parameters

- **table**: The code of a DPM Report table.
- **variable**: the code of a DPM *Variable*.
- **operation**: The code of a DPM *Operation*.
- **row, row_init, row_end**: A Reference to the code of a row of a DPM table.
- **col, col_init, col_end**: A Reference to the code of a col of a DPM table.
- **sheet, sheet_init, sheet_end**: A Reference to the code of a sheet of a DPM table.
- **interval_arithmetics**: A Boolean value specifying whether interval arithmetics should apply. Only applicable to numeric types (Number, Integer). When true, numeric values are treated as intervals with centre equal to the value and radius calculated from the precision of the number. The default value, in case the parameter is omitted, is false.
- **default_value**: A *Scalar* value specifying the value to be used in case of nulls. The value must be compatible with the *Data Type* of the selected *Variable*. The default value, in case the parameter is omitted, is null.

### 3.1.3 Output

> rset<*>

### 3.1.4 Semantics

Serves to select data as defined in the DPM model.

### 3.1.5 Additional Constraints

1. Existence of the parameters. The referred table, *Variable* or *Operation* need to be defined in the DPM.
2. *Operations* need to be defined in the same script.
3. Rows, columns and sheets need to exist in the table or *Operation* on which they are defined.
4. When the selection is done at *Variable* level, row, column and sheet parameters cannot exist.
5. The interval parameter can only be set to true when the selected *Variable* has a numeric *Data Type* (Number or Integer).

### 3.1.6 Behaviour

Returns a *Recordset* with the following structure:

*Key Components*:

- DPM-XL
  - Row: If the selected table has ordinate rows and more than one row is present in the selection. The name of the *Component* is "r".
  - Column: If the selected table has ordinate columns and more than one column is present in the selection. The name of the *Component* is "c".
  - Sheet: If the selected table has ordinate sheets and more than one sheet is present in the selection. The name of the *Component* is "s".

- DPM-ML
  - X: If there is Row for DPM-XL.
  - Y: If there is Column for DPM-XL.
  - Z: If there is Sheet for DPM-XL.

- A DPM *Key Component* for each *Key Variable* belonging to the *Key* associated to the selected *Variable*. The name of the *Component* is the *Code* of the Property associated to the *Key Variable*. The *Data Type* of the *Component* shall be the correspondent *Data Type* of the Metric or Property associated to the respective *Key Variable*.

*Fact Component*: With the *Data Type* corresponding to the selected *Variable*.

*Attribute Components*: One *Component* for each *Attribute Variable* associated to the selected *Variable*. The name of the *Component* is the *Code* of the Property associated to the *Attribute Variable*. The *Data Type* of the *Component* shall be the correspondent *Data Type* of the Metric or Property associated to the respective *Attribute Variable*.

The *Records* for the *Recordset* shall be obtained from the input data according to the selection.

Ranges (e.g., `r0010-0100`) select all *Headers* whose codes fall between the start and end codes, inclusive on both boundaries. The ordering is lexicographic, based on the string representation of the *Header* codes. In the example, all *Headers* with a code lexicographically between `0010` and `0100` are selected.

When a table has a *Header* that is not open, the selection needs to have an explicit value for that *Header*. The wildcard `*` (e.g., `r*`, `c*`, `s*`) is used to select all *Headers* of the corresponding dimension.

When an *Operation* is selected (using the `o` prefix), the resulting *Recordset* has the same structure as the output of the referenced *Operation*. The *Operation* must be defined in the same script, either as a named *Operation* in the DPM or via a non-persistent assignment (`:=`). Row, column, and sheet parameters can be used to further refine the selection on the *Operation* result.

#### 3.1.6.1 Example: Using an Operation as Input

```
op1 := {tT1, r010-020, c010} * 10;
{tO1, r010, c010} <- {oOp1, r010} + {tT1, r010, c020}
```

In this example, `op1` is assigned the result of multiplying the selection from table T1 by 10. The second statement selects row 010 from the result of `op1` (using the `o` prefix) and adds it to a selection from table T1. The structure of `{oOp1, r010}` is determined by the output of `op1`, filtered by row 010.

### 3.1.7 DPM-ML Metamodel Representation

There is no metamodel representation for the selection *Operator*. It is used to select, in last instance, *Variables*. So the result of the selection are processed to get the relevant *Variables* for DPM-ML, as well as the _DefaultValue_ and _UseIntervalArithmetics_ attributes.

## 3.2 With

### 3.2.1 Syntax

> **with** partial_selection [where_expression]<sup>1</sup>: _expression_

### 3.2.2 Input Parameters

- **partial_selection**: A selection expression (see Selection *Operator*)
- **expression**: An expression including selection *Operators*
- **where_expression**: An expression with a [`where` operator](./11-clause-operators.md#111-where)

### 3.2.3 Output

Does not generate output; modifies the selections of the expression after the colon.

### 3.2.4 Semantics

Serves to simplify expressions by adding a single context that may apply to all the [selection operators](#31-selection-operator) in the expression.

### 3.2.5 Additional Constraints

In case *where_expression* is used:

- All the operands rising from selection operators have a `where` or a `sub` clause applicable, the one in the with, or the one overriding it. This implies that it is not possible to have a selection with a table that does not have open keys.

### 3.2.6 Behaviour

The selection parameters in the partial selection applies to all the selections in the expression, unless they are overridden by a more specific parameter.

Therefore, in a selection inside the expression, whenever one of the parameters of the selection expression is not present, but that parameter is present on the partial selection, the parameter of the partial selection applies.

In case *where_expression* is used, the overriding mechanism works as follows:

- Any `where` clause used inside the expression fully overrides the *where_expression*, even if the properties used in the `where` clause inside the expression are different.
- A `sub` clause used inside the expression fully overrides the *where_expression*.

### 3.2.7 Examples

#### Example 1

```
with {tF_01.01, c0010, default: 0, interval: false}:
{r0010} = {r0020} + {r0030} + {r0040}
```

The partial selection applies to all the selections in the expression

#### Example 2

```
with {tF_01.01, c0010, default: 0, interval: false}:
{r0010} + {r0040} = {tF_04.01, r0010, c0010}
```

The partial selection applies to the two selections that do not refer to a table. Regarding the other selection, the default and intervals apply

#### Example 3

```
with {tF_01.01, c0010, default: 0, interval: false}:
{tF_01.01, r0010} + {tF_01.01, r0040} = {tF_04.01, r0010, c0010, default: null}
```

The partial selection applies to the two selections that refer to the same table referred to in the partial selection. Regarding the other selection, interval argument in the partial selection applies, but the default is overridden.

#### Example 4

```
with {c0010, default: 0, interval: false}:
{tF_01.01, r0010} + {tF_01.01, r0040} = {tF_04.01, r0010}
```

The partial selection applies to all the selections in the expression.

#### Example 5

```
with {tEXM, r0010}[where id1="ABC" and id2=5]:
  {c0010} = {c0020} + {c0030}
```

Where:

- **tEXM** has two *Key Components*: *id1* and *id2*

The clauses applying to each selection are:

| Selection   | Clause                       |
|-------------|------------------------------|
|`{c0010}`    |`[where id1="ABC" and id2=5]` |
|`{c0020}`    |`[where id1="ABC" and id2=5]` |
|`{c0030}`    |`[where id1="ABC" and id2=5]` |

#### Example 6

```
with {tEXM, r0010}[where id1="ABC" and id2=5]:
  {c0010} = {c0020} + {tEXM_02, c0030}[where id1="ABC"]
```


Where:

- **tEXM** has two *Key Components*: *id1* and *id2*
- **tEXM_02** has one *Key Component2*: *id1*

The clauses applying to each selection are:

| Selection   | Clause                       |
|-------------|------------------------------|
|`{c0010}`    |`[where id1="ABC" and id2=5]` |
|`{c0020}`    |`[where id1="ABC" and id2=5]` |
|`{tEXM_02, c0030}`|    `[where id1="ABC"]`  |


#### Example 7

```
with {r0010}[where id1="ABC"]:
  {tEXM, c0010} = {tEXM, c0020} + {tEXM_02, c0030}
```

Where:

- **tEXM** has two *Key Components*: *id1* and *id2*
- **tEXM_02** has one *Key Component*: *id1*

The clauses applying to each selection are:

| Selection         | Clause              |
|-------------------|---------------------|
|`{tEXM, c0010}`    |`[where id1="ABC"]`  |
|`{tEXM, c0020}`    |`[where id1="ABC"]`  |
|`{tEXM_02, c0030}` |`[where id1="ABC"]`  |


#### Example 8

```
with {default: 0, interval: true}[where qEEA = [eba_qAE:qx2010]]:
if
    sum({tC_08.01.a, r0070, c0255}) 
       =
    sum({tC_08.01.a, r0010, c0255})
then
    sum
       ( filter ({tC_08.02, c0255}, {tC_08.02, c0010} = 1)) 
        =
    sum({tC_09.02, r0090, c0120} [where CEG = [eba_GA:qx2014]]) 
endif
```

Where:

- **C_08.01.a** has one *Key Component*: *qEEA*
- **C_08.02** has two *Key Components*: *qEEA* and *OGR*
- **C_09.02** has one*Key Component*: *CEG*

The clauses applying to each selection are:

| Selection                  | Clause                           |
|----------------------------|----------------------------------|
|`{tC_08.01.a, r0070, c0255}`|`[where qEEA = [eba_qAE:qx2010]]` |
|`{tC_08.01.a, r0010, c0255}`|`[where qEEA = [eba_qAE:qx2010]]` |
|`{tC_08.02, c0255}`         |`[where qEEA = [eba_qAE:qx2010]]` |
|`{tC_08.02, c0010}`         |`[where qEEA = [eba_qAE:qx2010]]` |
|`{tC_09.02, r0090, c0120}`  |`[where CEG = [eba_GA:qx2014]]`   |


#### Example 9

```
with {tEXM, r0010}[where id1="ABC" and id2=5]:
  {c0010} = {c0020} + {tEXM_02, c0030}[sub id1="ABC"]
```

Where:

- **tEXM** has two *Key Components*: *id1* and *id2*
- **tEXM_02** has one *Key Component2*: *id1*

The clauses applying to each selection are:

| Selection         | Clause                      |
|------------------|------------------------------|
|`{c0010}`         |`[where id1="ABC" and id2=5]` |
|`{c0020}`         |`[where id1="ABC" and id2=5]` |
|`{tEXM_02, c0030}`|`[sub id1="ABC"]`             |

Note that the `sub` clause overrides the `where` in the `with` in the partial_selection.

#### Example 10

```
with {tEXM, r0010}[where id1="ABC" and id2=5]:
  {c0010} = {c0020} + {tEXM_03, c0030}
```

Where:

- **tEXM** has two *Key Components*: *id1* and *id2*
- **tEXM_03** does not have any key component

This expression raises a *Semantic Error*, because the where clause applies to the operand `{tEXM_03, c0030}`, and this table doesnt have the referred *Key Components*.


### 3.2.8 Metamodel Representation

See selection *Operator*.
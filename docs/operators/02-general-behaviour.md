# 2 Common General Behaviours for the Operators

## 2.1 Unary Operators

This general behaviour is applied to *Operators* taking as an argument one single *Operand*.

Unless explicitly specified differently, the behaviour for the unary *Operators* is as follows:

### 2.1.1 Scalars

The *Operator* is applied on a *Scalar* value and returns a *Scalar* value.

### 2.1.2 Recordsets

The *Operator* is applied on a *Recordset* and returns a *Recordset*.

The *Operator* is applied to the values of all the facts of the *Recordset*.

#### 2.1.2.1 Constraints

1. The application of the *Operator* is only allowed if all the facts of the *Operand* *Recordset* are compatible with the *Operator*.

#### 2.1.2.2 Structure

The structure of the resulting *Recordset* contains the *Key* and *Fact Components* of the *Operand* *Recordset*, the *Attribute Components* are not propagated.

#### 2.1.2.3 Records

The *Operators* produce one output *Record* per each input *Record*, which have:

- For *Key Components*, the values unchanged.
- For the *Fact* values, the *Operator* is applied to the input value and returns the corresponding value.

## 2.2 Binary Operators

This general behaviour is applied whenever an *Operator* takes as input two *Operands*, which includes the following *Operators*:

- [Arithmetic Binary Operators](./04-numeric-operators.md#43-binary-arithmetic-operators)
- [Comparison Operators](./05-comparison-operators.md#51-comparison-operators-general-behaviour) 
- [Logical Binary Operators](./06-logical-operators.md#6-logical)

Unless explicitly specified differently, the behaviour for the binary *Operators* is as follows.

### 2.2.1 Scalars

If the two *Operands* of a binary *Operator* are *Scalars*, the result shall be the *Scalar* resulting of applying the *Operator* to the *Operands*.

### 2.2.2 Recordset and Scalar

A binary *Operator* applied to a *Recordset* *Operand* and a *Scalar*, will result in a *Recordset* with the same *Key* and *Fact Components* as the input *Recordset* *Operand* (*Attribute Components* are not propagated). The *Operator* shall be applied to every *Record* of the input *Recordset* and the *Scalar*.

#### 2.2.2.1 Examples

```
0.25*{S.26.01, r0600, cNNN}
```

Supposing that the selection yields the following *Recordset*:

| c {#k} | f {#f} |
|--------|--------|
| 0060   | 100    |
| 0080   | 200    |

The operation results in:

| c {#k} | f {#f} |
|--------|--------|
| 0060   | 25     |
| 0080   | 50     |

### 2.2.3 Recordsets

#### 2.2.3.1 Constraints

Binary *Operators* can only be applied to two *Recordsets* *Operands* if they have:

1. exactly the same *Key Components*; or
2. the *Key Components* of one *Recordset* (Reference *Recordset*) are a superset of the *Key Components* of the other *Recordset*.

#### 2.2.3.2 Structure

The *Operator* yields a *Recordset* with the common *Key Components* in case 1, or the *Key Components* of the Reference *Recordset* in case 2. The resulting *Recordset* does not contain any *Attribute Component* from the *Operand* *Recordsets*.

#### 2.2.3.3 Records

The *Operator* applies to the pairs of values resulting from performing an inner join of the input *Recordsets* on the common *Key Components*.

#### 2.2.3.4 Examples

##### 2.2.3.4.1 Same Identifiers

```
{C 28.00 c040} + {C 28.00 c190}
```

Supposing that the selections yield:

{C 28.00 c040}

| INC {#k} | f {#f} |
|----------|--------|
| 123      | 1000   |
| 456      | 2000   |
| 789      | 3000   |

{C 28.00 c190}

| INC {#k} | f {#f} |
|----------|--------|
| 123      | -100   |
| 456      | -200   |
| 789      | -300   |

The result would be:

| INC {#k} | f {#f} |
|----------|--------|
| 123      | 900    |
| 456      | 1800   |
| 789      | 2700   |

##### 2.2.3.4.2 Subset of Identifiers

```
{F 40.01 c0110} >= {F 40.02c0060}
```

Supposing that the selections yield:

{F 40.01 c0110}

| LIN {#k} | TYC {#k} | f {#f} |
|----------|----------|--------|
| 123      | x1       | 1      |
| 456      | x1       | 0.8    |
| 789      | x1       | 0.4    |

{F 40.02 c0060}

| LIN {#k} | TYC {#k} | STC {#k} | LHC {#k} | LHO {#k} | f {#f} |
|----------|----------|----------|----------|----------|--------|
| 123      | x1       | 111      | ABC      | x1       | 0.3    |
| 123      | x1       | 111      | DEF      | x1       | 0.7    |
| 456      | x1       | 222      | ABC      | x1       | 0.85   |

The result would be:

| LIN {#k} | TYC {#k} | STC {#k} | LHC {#k} | LHO {#k} | f {#f} |
|----------|----------|----------|----------|----------|--------|
| 123      | x1       | 111      | ABC      | x1       | true   |
| 123      | x1       | 111      | DEF      | x1       | true   |
| 456      | x1       | 222      | ABC      | x1       | false  |

Note that:

- An inner join on the common *Key Components* is performed. This means that there is no comparison for the *Record* in F 40.01 with LIN = 789, because there is no match in F 40.02.
- The result of the operation has the *Structure* of the Reference *Recordset*, which is the one that has a superset of *Key Components*.
- For each match in the join the *Operator* `>=` is applied to the pairs of values for the facts.

### 2.2.4 Intervals Treatment

When one of two numeric *Operands* is an interval, and the other is a point, the point is transformed into an interval with centre value equal to the point value, and radius 0.
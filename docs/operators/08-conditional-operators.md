# 8 Conditional Operators

## 8.1 If-then-else

### 8.1.1 Syntax

> **if** conditionExpression **then** thenExpression {**else** elseExpression} **endif**

### 8.1.2 Input Parameters

> conditionExpression: rset | scal <boo\>

> thenExpression: rset | scal <*>

> elseExpression: rset | scal <*>

### 8.1.3 Output

> rset | scal <*>

### 8.1.4 Semantics
The result combines values selected from the thenExpression where the conditionExpression evaluates to `true`, with values selected from the elseExpression where the conditionExpression evaluates to `false` or `null`. *Records* are matched between the conditionExpression and each branch by inner join. When the elseExpression is omitted, *Records* where the conditionExpression is not `true` are excluded from the result. The formal evaluation is given in §8.1.6.

### 8.1.5 Additional Constraints

The additional constraints ensure that the structure of the output is known at compile time and that no *Record* in the result can have a `null` value in any *Key Component*.

To the effects of these additional constraints, *Scalars* are treated as *Recordsets* with zero *Key Components*.

- **Data Type compatibility**: The thenExpression and elseExpression (when provided) must be of the same *Data Type*, or of *Data Types* that admit implicit casting per the casting rules in [§2.3.2](../information-model/02-information-model.md#232-casting). When implicit casting applies, the resulting *Data Type* is that common *Data Type*.
- **Structure compatibility with the condition**: The *Key Components* of the thenExpression must be compatible with those of the conditionExpression — that is, one set must be equal to, or a subset of, the other, following the same rule as for [binary *Operators* on two *Recordsets*](./02-general-behaviour.md#2231-constraints). If the elseExpression is provided, the same restriction applies to the elseExpression together with the conditionExpression.
- **Matched join structure**: When the elseExpression is provided, the *Key Components* of the thenExpression and the elseExpression must satisfy one of: a) they are the same; or b) both are equal to, or a subset of, those of the conditionExpression.
- **Null literal exception**: When thenExpression or elseExpression is `null` (explicit or, in the case of elseExpression, implicit because not used), it is possible for the other expression to have a set of key components that is a superset of the conditionExpression. This is allowed because null records are not added to the output, and therefore no records with null value for *Key components* may arise.

### 8.1.6 Behaviour

Throughout this section, `join(x, y)` denotes the inner join of [§2.2.3.3](./02-general-behaviour.md#2233-records); its result is a *Recordset* whose *Key Components* are the superset of those of x and y (which, under the *Structure compatibility with the condition* constraint, equals the *Key Components* of whichever operand is the superset). This makes the *Matched join structure* constraint of §8.1.5 equivalent to requiring that `join(conditionExpression, thenExpression)` and `join(conditionExpression, elseExpression)` have the same structure: case (a) gives both joins the *Key Components* `conditionExpression ∪ thenExpression = conditionExpression ∪ elseExpression`; case (b) gives both joins the *Key Components* of the conditionExpression.

- **Structure**:
    - If all operands are *Scalars* (zero *Key Components*), the result is a *Scalar*; otherwise the result is a *Recordset*.
    - The structure of the output shall contain the largest set of *Key Components* of the input structures.
- **Evaluation**: 
    - The expression is evaluated as the disjoint union of two sets of *Records*:
        - The *Records* of `join(conditionExpression, thenExpression)` for which the conditionExpression evaluates to `true`, with the value of the *Fact Component* taken from the thenExpression.
        - When the elseExpression is provided: the *Records* of `join(conditionExpression, elseExpression)` for which the conditionExpression evaluates to `false` or `null`, with the value taken from the elseExpression.
    - The two sets are disjoint by construction (they select complementary *Records* of the conditionExpression) and share the same structure by the *Matched join structure* constraint.
- **Omitted else**: *Records* where the conditionExpression evaluates to `false` or `null` are not included in the result.

The following table summarises the allowed combinations, the additional constraint that applies beyond pairwise structure compatibility, and the resulting structure:

To the effects of this table:

- An explicit `else null` is treated as a missing else.
- When thenExpression or elseExpression is the literal `null`, the *Null literal exception* of §8.1.5 may further relax the constraints in the corresponding row, since no *Records* are contributed by a null branch.

| Condition | Then | Else | Additional constraint | Result structure |
|-----------|------|------|-----------------------|------------------|
| *Scalar* | *Scalar* | — | — | *Scalar* |
| *Scalar* | *Scalar* | *Scalar* | — | *Scalar* |
| *Scalar* | *Recordset* | — | — | *Recordset* (keys of then) |
| *Scalar* | *Recordset* | *Recordset* | then and else must have the same *Key Components* | *Recordset* (common keys of then and else) |
| *Scalar* | *Scalar* | *Recordset* | Forbidden — join structures differ | — |
| *Scalar* | *Recordset* | *Scalar* | Forbidden — join structures differ | — |
| *Recordset* | *Scalar* | — | — | *Recordset* (keys of condition) |
| *Recordset* | *Scalar* | *Scalar* | — | *Recordset* (keys of condition) |
| *Recordset* | *Scalar* | *Recordset* | *Key Components* of else ⊆ those of condition | *Recordset* (keys of condition) |
| *Recordset* | *Recordset* | — | *Key Components* of then and condition: equal or one a subset of the other| *Recordset* (superset of keys) |
| *Recordset* | *Recordset* | *Scalar* | *Key Components* of then ⊆ those of condition | *Recordset* (keys of condition) |
| *Recordset* | *Recordset* | *Recordset* | `join(condition, then)` and `join(condition, else)` must share the same structure | *Recordset* (common join structure) |

**Relation to other operators**: Under this definition, `if-then-else` subsumes several other operators as equivalent idioms:

- `if A then B` (no else), when the *Key Components* of A are a subset of those of B, is equivalent to `filter(B, A)`.
- `if isnull(A) then B else A` is equivalent to `nvl(A, B)` when A and B share the same data structure (the case in which `nvl` is applicable).
- `if A = x then y else A` is a general pattern to replace a value `x` with `y`.

### 8.1.7 Examples

#### Example 1: Scalar condition, scalar branches

```
if 10 > 5 then 100 else 200 endif
```

Result: `100` (scalar). The condition is true, so the then branch is returned.

#### Example 2: Scalar condition, recordset branches

Supposing `{tT1, r010-020}` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 100    |
| 020    | 200    |

And `{tT2, r010-020}` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 900    |
| 020    | 800    |

```
if 10 > 5 then {tT1, r010-020} else {tT2, r010-020} endif
```

Result:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 100    |
| 020    | 200    |

The condition is true, so the entire then recordset is returned.

#### Example 3: Recordset condition, scalar branches

Supposing `{tT1, r010-020} > 150` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | false  |
| 020    | true   |

```
if {tT1, r010-020} > 150 then 1 else 0 endif
```

Result:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 0      |
| 020    | 1      |

The condition is evaluated per record. The scalar branches are applied to each record based on its condition value.

#### Example 4: Recordset condition, recordset branches with same keys

Supposing `{tT1, r010-020}` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 100    |
| 020    | 200    |

And `{tT2, r010-020}` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 900    |
| 020    | 800    |

```
if {tT1, r010-020} > 150 then {tT1, r010-020} else {tT2, r010-020} endif
```

Result:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 900    |
| 020    | 200    |

Record 010 takes the else value (100 is not > 150), record 020 takes the then value (200 > 150).

#### Example 5: Forbidden — scalar condition with mixed branches

The following expression is **not allowed**:

```
if {tT1, r010} > 150 then {tT2, r010-020} else 0 endif
```

The condition `{tT1, r010} > 150` is a *Scalar* Boolean. The then branch is a *Recordset* with *Key Components* `{r}` and the else branch is a *Scalar* (zero *Key Components*). Applying the *Matched join structure* constraint: `join(condition, then)` has keys `{r}` while `join(condition, else)` has no keys. Since these structures differ, the expression is rejected. To fix this, either omit the else branch (records where the condition is `false` or `null` will not appear in the result) or replace the else branch with a *Recordset* sharing the same *Key Components* as the then branch.

#### Example 6: Recordset condition with more keys than the branches

Supposing `{tT1, r010-020}` is an open table with a *Key Component* CNT, and the condition `{tT1, r010-020} > 0` yields:

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 010    | ES       | true   |
| 010    | PT       | false  |
| 020    | ES       | true   |
| 020    | PT       | true   |

And the then branch `{tT2, r010-020}` has only the r key:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 100    |
| 020    | 200    |

```
if {tT1, r010-020} > 0 then {tT2, r010-020} else 0 endif
```

The condition has keys {r, CNT} and the then branch has keys {r}. The result has the structure of the superset (the condition), and the then/else values are matched via inner join on the common key (r):

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 010    | ES       | 100    |
| 010    | PT       | 0      |
| 020    | ES       | 200    |
| 020    | PT       | 200    |

#### Example 7: Omitted else

Supposing `{tT1, r010-020}` yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 100    |
| 020    | 200    |

```
if {tT1, r010-020} > 150 then {tT1, r010-020} endif
```

Result:

| r {#k} | f {#f} |
|--------|--------|
| 020    | 200    |

Record `010` is excluded because its condition value is `false` (100 is not > 150). Record `020` is included with its then value.

#### Example 8: Forbidden — branches with incompatible key structures

Supposing `{tT1, r010-020}` has *Key Component* `r` and yields:

| r {#k} | f {#f} |
|--------|--------|
| 010    | 200    |
| 020    | 100    |

And `{tT3, r010-020}` has *Key Components* `r` and `CNT` and yields:

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 020    | ES       | 500    |
| 020    | PT       | 600    |

The following expression is **not allowed**:

```
if {tT1, r010-020} > 150 then 0 else {tT3, r020} endif
```

`join(condition, then)` has keys `{r}` while `join(condition, else)` has keys `{r, CNT}`. Since the structures differ, the *Matched join structure* constraint rejects the expression at compile time. This prevents results where records of the condition would silently disappear because one branch cannot supply values for the key tuples required by the other (while the same records would be included if the elseExpression was omitted). To fix this, ensure both branches produce results with the same *Key Components* after joining with the condition.

#### Example 9: Null literal exception — then branch with more keys than the condition

Supposing the condition `{tT1, r010-020} > 0` yields a *Recordset* with key `{r}`:

| r {#k} | f {#f} |
|--------|--------|
| 010    | true   |
| 020    | false  |

And the then branch `{tT2, r010-020}` is an open table with keys `{r, CNT}`:

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 010    | ES       | 100    |
| 010    | PT       | 200    |
| 020    | ES       | 300    |
| 020    | PT       | 400    |

```
if {tT1, r010-020} > 0 then {tT2, r010-020} endif
```

With no `else` clause, the elseExpression is implicitly the literal `null`. The *Matched join structure* constraint would normally reject this — `join(condition, then)` has *Key Components* `{r, CNT}` while `join(condition, null)` has `{r}` — but the *Null literal exception* of §8.1.5 admits the case because the implicit-null branch contributes no *Records*, so no *Record* with a `null` value in a *Key Component* can arise.

Result:

| r {#k} | CNT {#k} | f {#f} |
|--------|----------|--------|
| 010    | ES       | 100    |
| 010    | PT       | 200    |

The result takes the structure of the then branch, `{r, CNT}`. *Records* `(010, ES)` and `(010, PT)` are included (the condition is `true` for r = 010); *Records* with r = 020 are excluded because the condition is `false` and the implicit-null else branch contributes nothing.

## 8.2 Null Substitute (nvl)

### 8.2.1 Syntax

> **nvl(** op1, op2 **)**

### 8.2.2 Input Parameters

> op1: rset | scal <*>

> op2: rset | scal <*>

### 8.2.3 Output

> rset | scal <*>

### 8.2.4 Semantics
Returns op2 when op1 is `null`, otherwise op1.

### 8.2.5 Additional Constraints
op1 and op2 need to be of the same type (*Scalar* or *Recordset*) and same *Data Type*.

If op1 and op2 are of the *Recordset* type, they need to have the same data structure.

### 8.2.6 Behaviour

[Binary *Operators'* standard behaviour](./02-general-behaviour.md#22-binary-operators).

### 8.2.7 Examples

## 8.3 Filter

### 8.3.1 Syntax

> **filter(** filteredOp, filteringOp **)**

### 8.3.2 Input Parameters

> filteredOp: rset <*>

> filteringOp: rset <boo\>

### 8.3.3 Output

> rset <*>

### 8.3.4 Semantics
Returns the *Records* of the filteredOp *Recordset* for which their matching *Record* in filteringOp evaluates to `true`. *Records* whose matching value in filteringOp is `false` or `null` are excluded (per the null-as-false rule for filtering in [§4.3](../information-model/04-null-treatment.md#43-null-treatment)).

### 8.3.5 Additional Constraints
filteringOp must have the same *Key Components* as filteredOp or a subset of them.

### 8.3.6 Behaviour
Returns a *Recordset* with the same data structure as filteredOp.

An inner join between filteredOp and filteringOp is performed. The *Records* in filteredOp that match to a *Record* in filteringOp with true value are returned.

### 8.3.7 Examples
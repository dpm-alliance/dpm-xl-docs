# 11 Clause Operators

Clause *Operators* serve to perform operations on the *Components* of *Recordsets*. Which roles of *Component* (*Key*, *Fact* or *Attribute*) each clause *Operator* accepts is specified in its Additional Constraints: `where` may reference any of them, while `rename`, `get` and `sub` apply only to *Key* and, where indicated, *Attribute Components*. Use of clause *Operators* with Standard *Key Components* ("r", "c" or "s") is not allowed.

## 11.1 where

### 11.1.1 Syntax

> op **[where** condition **]**

### 11.1.2 Input Parameters

> op: rset <*>

> condition: expression <boo\>

### 11.1.3 Output

> rset <*>

### 11.1.4 Semantics
Filters a *Recordset* based on the values of its *Components*.

### 11.1.5 Additional Constraints
Condition must be a Boolean expression using as input *Components* of op — any of its DPM *Key Components*, its *Fact Component*, or its *Attribute Components*. The contextual *Components* are latent in every *Recordset* and may therefore always be referenced: `refPeriod` and `entityID` ([§2.2.3.1.4](../information-model/02-information-model.md#22314-contextual-key-components)), and the contextual *Attribute Components* ([§2.2.3.3.1](../information-model/02-information-model.md#22331-contextual-attribute-components)).

Standard *Key Components* ("r", "c" or "s") cannot be referenced in the condition.

### 11.1.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*, and with the *Records* resulting from the evaluation of the filtering condition. When the condition evaluates to true, the *Record* is kept, otherwise (including null), the *Record* is not kept.

### 11.1.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
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

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |

`{tT1, r010-020}[where f != 0]` keeps only the *Records* with a non-zero *Fact* value. In the example above, all the *Records* are kept, as none of them has a zero or null *Fact*.

`{tT1, r010-020}[where f > 300 and CNT = "PT"]` combines *Fact* and *Key* filtering, and returns:

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q2       | 020    | PT       | 600    |

`{tT1, r010-020}[where not isnull(f)]` keeps only the *Records* with a non-null *Fact* value.

## 11.2 rename

### 11.2.1 Syntax

> op **[rename** compFrom to compTo {, compFrom to compTo}* **]**

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

compFrom may refer to a DPM *Key Component* or to an *Attribute Component*, including a contextual *Attribute Component* ([§2.2.3.3.1](../information-model/02-information-model.md#22331-contextual-attribute-components)), which is latent in every *Recordset* and therefore always belongs to op. It cannot refer to a Standard *Key Component* (i.e., cannot refer to row, column or sheet), nor to the *Fact Component* (the name of the *Fact Component* is always "f", as specified in §2.2.3.2).

compTo name cannot be already used in op.

compTo name cannot be the name of a standard *Key Component* ("r", "c" or "s"), nor the name of a contextual *Component* ([§2.4.3](../information-model/02-information-model.md#243-reserved-words)), which would make the contextual *Component* unaddressable.

### 11.2.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*, except the renamed *Components*, which change name. The role of each renamed *Component* (*Key* or *Attribute*) is preserved. All the *Records* are kept.

### 11.2.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
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

| refPeriod {#k} | r {#k} | country {#k} | f {#f} |
|--------------|--------|--------------|--------|
| 2022Q1       | 010    | ES           | 100    |
| 2022Q1       | 010    | PT           | 200    |
| 2022Q1       | 010    | DE           | 300    |
| 2022Q1       | 010    | IT           | 400    |
| 2022Q2       | 020    | ES           | 500    |
| 2022Q2       | 020    | PT           | 600    |
| 2022Q2       | 020    | DE           | 700    |
| 2022Q2       | 020    |              | 800    |

Considering now a *Recordset* with an *Attribute Component* "p" (precision), generated from the selection `{tT2, r010}`:

| refPeriod {#k} | CNT {#k} | f {#f} | p {#a} |
|--------------|----------|--------|--------|
| 2022Q1       | ES       | 100    | 1      |
| 2022Q1       | PT       | 200    | 0      |

`{tT2, r010}[rename p to precision]` renames the *Attribute Component*, and returns:

| refPeriod {#k} | CNT {#k} | f {#f} | precision {#a} |
|--------------|----------|--------|----------------|
| 2022Q1       | ES       | 100    | 1              |
| 2022Q1       | PT       | 200    | 0              |

## 11.3 get

### 11.3.1 Syntax

> op **[get** component **]**

### 11.3.2 Input Parameters

> op: rset <*>

> component: scal <prop\>

### 11.3.3 Output

> rset <*>

### 11.3.4 Semantics
Returns a *Recordset* where the *Fact* is the value of one of the *Key* or *Attribute Components*.

### 11.3.5 Additional Constraints
component must belong to op.

component may refer to a DPM *Key Component* or to an *Attribute Component*, including a contextual *Attribute Component* ([§2.2.3.3.1](../information-model/02-information-model.md#22331-contextual-attribute-components)), which is latent in every *Recordset* and therefore always belongs to op. It cannot refer to a Standard *Key Component* (i.e., cannot refer to row, column or sheet), nor to the *Fact Component* (the operation would have no effect).

### 11.3.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*. The *Fact* values the *Records* are substituted for the value of the selected *Component*. The *Data Type* of the *Fact Component* of the result is the *Data Type* of the selected *Component*.

*Records* with null *Fact* value will be omitted.

### 11.3.7 Examples
Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
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

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | ES     |
| 2022Q1       | 010    | PT       | PT     |
| 2022Q1       | 010    | DE       | DE     |
| 2022Q1       | 010    | IT       | IT     |
| 2022Q2       | 020    | ES       | ES     |
| 2022Q2       | 020    | PT       | PT     |
| 2022Q2       | 020    | DE       | DE     |
| 2022Q2       | 020    | IT       | IT     |

Considering now a *Recordset* with an *Attribute Component* "p" (precision), generated from the selection `{tT2, r010}`:

| refPeriod {#k} | CNT {#k} | f {#f} | p {#a} |
|--------------|----------|--------|--------|
| 2022Q1       | ES       | 100    | 1      |
| 2022Q1       | PT       | 200    | 0      |

`{tT2, r010}[get p]` promotes the value of the *Attribute Component* into the *Fact Component*, so that it can be further aggregated or compared, and returns:

| refPeriod {#k} | CNT {#k} | f {#f} | p {#a} |
|--------------|----------|--------|--------|
| 2022Q1       | ES       | 1      | 1      |
| 2022Q1       | PT       | 0      | 0      |

## 11.4 sub 

### 11.4.1 Syntax

> op **[sub** key_component = value {, key_component = value}* **]**

### 11.4.2 Input Parameters

> op: rset <*>

> key_component: scal <prop\>

The right-hand side `value` accepts three forms:

| Form | Syntax | Description |
|------|--------|-------------|
| Literal | A scalar literal value | A constant of a basic type (String, Number, Boolean, Date, etc.) written directly in the expression. |
| Select | `{` selectOperand `}` | A selection expression that resolves to a *Scalar* value (e.g., a single-cell table selection or a *Variable* reference). |
| Item reference | `[` dimension `:` member `]` | A dimension member reference using the `dimension:member` notation. |

### 11.4.3 Output

> rset | scal <*>

### 11.4.4 Semantics
Filters a *Recordset* based on the value of a *Key Component*, and then drops the *Key Component*. If all *Key Components* are substituted, the result is a *Scalar*.

### 11.4.5 Additional Constraints
key_component must belong to op, and must be a DPM *Key Component*. It cannot refer to the *Fact Component* (substituting it would leave the *Recordset* without a value) nor to an *Attribute Component*.

The value needs to have the same *Data type* as the key_component.

### 11.4.6 Behaviour
Returns a *Recordset* with the same data structure as the input *Operand*, except the *Key Components* dropped, and with the *Records* resulting from the evaluation of the filtering condition. When the condition evaluates to true, the *Record* is kept, otherwise (including null), the *Record* is not kept. If no *Key Components* remain after the substitution, the result is a *Scalar*.

### 11.4.7 Examples

Considering the following *Recordset*, generated from the selection `{tT1, r010-020}`:

| refPeriod {#k} | r {#k} | CNT {#k} | f {#f} |
|--------------|--------|----------|--------|
| 2022Q1       | 010    | ES       | 100    |
| 2022Q1       | 010    | PT       | 200    |
| 2022Q1       | 010    | DE       | 300    |
| 2022Q1       | 010    | IT       | 400    |
| 2022Q2       | 020    | ES       | 500    |
| 2022Q2       | 020    | PT       | 600    |
| 2022Q2       | 020    | DE       | 700    |
| 2022Q2       | 020    |          | 800    |

**Form 1 — Literal**

`{tT1, r010-020}[sub CNT = "PT"]` returns:

| refPeriod {#k} | r {#k} | f {#f} |
|--------------|--------|--------|
| 2022Q1       | 010    | 200    |
| 2022Q2       | 020    | 600    |

**Form 2 — Select**

Where `{vCountry}` is a *Variable* that resolves to the String `"ES"`:

`{tT1, r010-020}[sub CNT = {vCountry}]` returns:

| refPeriod {#k} | r {#k} | f {#f} |
|--------------|--------|--------|
| 2022Q1       | 010    | 100    |
| 2022Q2       | 020    | 500    |

**Form 3 — Item reference**

`{tT1, r010-020}[sub CNT = [CNT:ES]]` returns:

| refPeriod {#k} | r {#k} | f {#f} |
|--------------|--------|--------|
| 2022Q1       | 010    | 100    |
| 2022Q2       | 020    | 500    |

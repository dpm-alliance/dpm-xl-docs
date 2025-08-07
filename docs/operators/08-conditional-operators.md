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
Returns the thenExpression if the conditionExpression evaluates to `true`, elseExpression otherwise.

### 8.1.5 Additional Constraints
The thenExpression and elseExpression must be of the same *Data Type* or a compatible one. In case they are not of the same *Data Type*, there shall be implicit casting to the common *Data Type*, which will be the *Data Type* of the result

The thenExpression and elseExpression need to have the same data structure.

If the conditionExpression is of *Recordset* type, the data structures of the thenExpression and elseExpression must contain the same identifiers as the conditionExpression, or a subset of them. The resulting structure will contain the *Components* of the conditionExpression *Recordset*.

The following table contains the applicability options for the if-then-else *Operator*:

| Condition {#k} | Then {#k} | Else {#k} | Is allowed {#k} | Result structure {#f} | Remarks {#f} |
|----------------|-----------|-----------|-----------------|----------------------|--------------|
| *Scalar*         | *Scalar*    | None      | Yes             | *Scalar*               |              |
| *Scalar*         | *Recordset* | None      | No              | -                    | Forbidden because the result structure would be unknown |
| *Scalar*         | *Scalar*    | *Scalar*    | Yes             | *Scalar*               |              |
| *Scalar*         | *Scalar*    | *Recordset* | No              | -                    | Forbidden because the result structure would be unknown |
| *Scalar*         | *Recordset* | *Scalar*    | No              | -                    | Forbidden because the result structure would be unknown |
| *Scalar*         | *Recordset* | *Recordset* | Yes             | *Recordset*            |              |
| *Recordset*      | *Scalar*    | None      | Yes             | *Recordset*            |              |
| *Recordset*      | *Recordset* | None      | Yes             | *Recordset*            |              |
| *Recordset*      | *Scalar*    | *Scalar*    | Yes             | *Recordset*            |              |
| *Recordset*      | *Scalar*    | *Recordset* | Yes             | *Recordset*            |              |
| *Recordset*      | *Recordset* | *Scalar*    | Yes             | *Recordset*            |              |
| *Recordset*      | *Recordset* | *Recordset* | Yes             | *Recordset*            |              |

### 8.1.6 Behaviour
For conditionExpressions of the *Scalar* type, returns the *Operand* resulting from the thenExpression or the elseExpression.

For conditionExpressions of the *Recordset* type, returns a *Recordset* with the data structure of the conditionExpression, with one output *Record* per input *Record*. The *Fact* values will be the corresponding for the thenExpression or elseExpression, depending on the evaluation result of the condition.

If the elseExpression is omitted, a null is returned, except for the case when the thenExpression is of the Boolean *Data Type*, in which case true is returned.

If the condition evaluates to null, the elseExpression is returned.

### 8.1.7 Examples

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

> filteringOp: rset <bool\>

### 8.3.3 Output

> rset <*>

### 8.3.4 Semantics
Returns the *Records* of the filteredOp *Recordset* for which their matching *Record* in filteringOp evaluates to true.

### 8.3.5 Additional Constraints
filteringOp must have the same *Key Components* as filteredOp or a subset of them.

### 8.3.6 Behaviour
Returns a *Recordset* with the same data structure as filteredOp.

An inner join between filteredOp and filteringOp is performed. The *Records* in filteredOp that match to a *Record* in filteringOp with true value are returned.

### 8.3.7 Examples
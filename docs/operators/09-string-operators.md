# 9 String Operators

## 9.1 Length (len)

### 9.1.1 Syntax

> **len(** op **)**

### 9.1.2 Input Parameters

> op: rset | scal <str\>

### 9.1.3 Output

> rset | scal <int\>

### 9.1.4 Semantics
Returns the number of characters of the op string.

### 9.1.5 Additional Constraints
None.

### 9.1.6 Behaviour

- [Unary *Operators'* standard behaviour](./02-general-behaviour.md#21-unary-operators).

### 9.1.7 Examples

- `len("test")` results in `4`
- `len(null)` results in `null`

## 9.2 Concatenate (&)

### 9.2.1 Syntax

> op1 **&** op2

### 9.2.2 Input Parameters

> op1: rset | scal <str\>

> op2: rset | scal <str\>

### 9.2.3 Output

> rset | scal <str\>

### 9.2.4 Semantics
Concatenates two strings.

### 9.2.5 Additional Constraints
None.

### 9.2.6 Behaviour
- [Binary *Operators'* standard behaviour](./02-general-behaviour.md#22-binary-operators).

### 9.2.7 Examples
- `"hello" & "world"` results in `"helloworld"`
- `"test" & null` results in `null`
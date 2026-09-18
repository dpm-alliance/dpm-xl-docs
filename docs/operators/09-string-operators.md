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

## 9.3 Substring (substr)

### 9.3.1 Syntax

> **substr(** op, {start}, {length} **)**

### 9.3.2 Input Parameters

> op: rset | scal <str\>

> start: scal <int\>

> length: scal <int\>

### 9.3.3 Output

> rset | scal <str\>

### 9.3.4 Semantics
Extracts the substring of op that begins at the character in position start (the first character of op is in position 1) and runs for length characters.

start and length are optional. If start is omitted, extraction begins at the first character of op. If length is omitted, the substring runs to the end of op. If both are omitted, the whole op string is returned.

### 9.3.5 Additional Constraints
- start must be greater than or equal to 1.
- length must be greater than or equal to 0. A length of 0 is permitted (and returns an empty string) so that a length derived from another expression, such as the difference between two positions, can collapse to an empty span without raising an error.

### 9.3.6 Behaviour
- [Unary *Operators'* standard behaviour](./02-general-behaviour.md#21-unary-operators).
- If start is greater than the number of characters of op, an empty string is returned.
- If start + length goes past the end of op, the substring ends at the last character of op.

### 9.3.7 Examples
- `substr("Hello World", 7)` results in `"World"`
- `substr("Hello World", 1, 5)` results in `"Hello"`
- `substr("Hello World", 7, 100)` results in `"World"`
- `substr("Hello World")` results in `"Hello World"`
- `substr("Hello World", 20)` results in `""`
- `substr("Hello World", 7, 0)` results in `""` (a length of 0 extracts no characters)
- `substr(null, 2, 3)` results in `null`
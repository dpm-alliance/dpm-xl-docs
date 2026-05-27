# 12 Assignment Operators

Assignment *Operators* serve to assign the result of an operation to a name or to a *DPM Variable*.

## 12.1 Non-persistent assignment (:=)

### 12.1.1 Syntax

> name **:=** op

### 12.1.2 Input Parameters

> name: the name assigned to the result

> op: the operand, which can be any dpm-xl expression

### 12.1.3 Output

> Assigns the name to the op.

### 12.1.4 Semantics
Allows a *DPM-XL* operation to use the result of another operation as input. By assigning a name to the result of an operation, this result can be selected using the selection operator.

### 12.1.5 Additional Constraints
name must be a valid temporary identifier as defined in the grammar.

### 12.1.6 Behaviour
The value of the *Operand* op is assigned the name *name*, which is non-persistent and therefore is not stored. The operand op can be obtained through an expression as complex as needed (for example: `op1 := {tT1, r010-020} * 10`).

## 12.2 Persistent assignment (<-)

### 12.2.1 Syntax

> var_sel **<-** op

### 12.2.2 Input Parameters

> var_sel: a selection operator with the output variable

> op: the operand, which can be any dpm-xl expression

### 12.2.3 Output

> Assigns the result of op to the var_sel.

### 12.2.4 Semantics

Links the calculation of a *DPM Variable* to an expression.

### 12.2.5 Additional Constraints

var_sel is a selection *Operator* that refers to a *Variable* or set of *Variables* that are defined in the *DPM*.
The resulting *Structure* in op needs to be equal to the *Structure* of var_sel.


### 12.2.6 Behaviour
The value of the *Operand* op is assigned the *Variables* selected in var_sel, with the objective of storing the result. The operand op can be obtained through an expression as complex as needed (for example: `{tOT1, r010-020} <- {tT1, r010-020} * 10`).


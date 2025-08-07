# 4 Null in DPM Expression Language

## 4.1 Null and Data Types

All the *Data Types* are assumed to contain the conventional value null, which means "no value", or "absence of known value" or "missing value". Note that the null value, therefore, is the only value of multiple different types.

For the String *Data Type*, null is considered equivalent to the empty string.

## 4.2 Implicit and Explicit Null Values

For *Recordsets*, nulls may arise in a selection for two reasons:

- **Explicit null**: In the input data for the engine, there is a *Record* without a value.
- **Implicit null**: In the input data for the engine, there is no *Record* for one of the *Variables* in the selection. In this case, the engine will generate a *Record* for that *Variable* with a null value, or another specified default (see Selection *Operator* in part 2).

In any case, there are no differences in how implicit or explicit nulls are treated.

## 4.3 Null Treatment

In general, most of the *Operations* return null when any of their arguments is null.

- **Comparison Operators** (e.g., -, >): if a null is involved in the operation, then the result is null.
- **Arithmetic Operators** (e.g., + , -, *): if a null is involved in the operation, then the result is null.
- **String Operators**: null is considered an empty string.
- **Logical Operators** (and, or, xor, not): Three-value logic is adopted with the consideration that null means unknown. The concrete results for each *Operator* are specified in part 2, in the description of the *Operators*.
- **Conditional Operators**: null is considered equivalent to false.
- **Filtering**: null is considered equivalent to false (i.e., the *Records* with a null value are not selected in the filter).
- **Aggregations** (e.g., sum, avg, max): nulls are excluded from the calculations.
- **Intervals**: If the centre is null, the radius is also null.

## 4.4 Overriding the Standard Null Treatment

The standard null treatment can be overridden by substituting the null values for other values. This can be done in two ways:

- **With the Selection Operator**: It is possible to set a default value in the selection for the null values.
- **With the nvl Operator**: serves to substitute nulls for any value.
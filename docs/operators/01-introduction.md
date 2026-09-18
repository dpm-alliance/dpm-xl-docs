# 1 Introduction

This documentation provides the individual description of the semantics for all the *Operators* in the DPM Operations (DPM-XL and DPM-ML).

The document is structured from more generic to more concrete. It starts with general behaviours that apply to several *Operators* of different types. Then it follows with the common behaviour of the *Operators* of the same type, finalizing with the specific behaviour of the single *Operators*.

## 1.1 Operator Precedence

The table below lists operator precedence from lowest (level 1) to highest (level 12), derived from the order of alternatives in the grammar's `expression` rule. Operators at higher levels bind more tightly. When operators share a level they associate left-to-right. Use parentheses `()` to override any precedence.

| Level | Operator(s) | Category |
|------:|-------------|----------|
| 1 | `or`, `xor` | Logical disjunction |
| 2 | `and` | Logical conjunction |
| 3 | `in` | Set membership |
| 4 | `=`, `<>`, `<`, `>`, `<=`, `>=` | Comparison |
| 5 | `\|\|` | String concatenation |
| 6 | `+`, `-` | Additive |
| 7 | `*`, `/` | Multiplicative |
| 8 | `not(...)` | Logical negation |
| 9 | `+`, `-` (unary) | Unary sign |
| 10 | `[where ...]`, `[sub ...]`, `[rename ...]`, `[get ...]` | Clause operators (postfix) |
| 11 | Function calls, literals, selections, item references | Atomic expressions |
| 12 | `(...)` | Parenthesised grouping |

`if … then … else … endif` encloses a full expression and is always evaluated as a unit; its precedence is that of the enclosed expressions.

This ordering mirrors the precedence structure of VTL (Validation and Transformation Language), which DPM-XL is based on.
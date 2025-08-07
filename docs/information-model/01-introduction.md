# Introduction

This is a technical document specifying the Information Model, the *Data Types* and other technical characteristics of the DPM *Operations*.

This documentation is mainly addressed to technical audiences with interest in developing engines or in having a better technical understanding of the DPM *Operations*. The document is not addressed to normal business users that write calculations, although it can be useful for those users with basic knowledge of information models and programming languages.

In any case, this document is complemented by the document describing the *Operators* used for the *Operations*, which should be the main reference for users writing calculations.

## Why

EBA and EIOPA have been using for some years a semi-formal expression language for business users to write and share their data validation requirements.

The fact that is semi-formal allows automating some aspects for translating the expression into another language (e.g., XBRL Formula Link base). But because it is not completely formal, it is, by definition, not possible to create a compiler that fully automates the translation to another language or an interpreter for it.

This document, and its complementary document on the *Operators*, aims to provide this required formalization so that DPM *Operations* become fully formalized, therefore enabling the automation regarding the calculations written by EBA and EIOPA.

## DPM-XL and DPM-ML

The DPM *Operations* have two representations: As expressions, with the DPM eXpression Language (DPM-XL) and as a structured representation in the database following a metamodel (DPM Metamodel Language or DPM-ML).

The reason for having **DPM-XL** is that it is the language business users write and can understand. Because it is the input of all the process, it is important to keep this language formal, so that it can be translated and executed. But, at the same time, it must serve for communication, so it has to be business users friendly, implying, among others, that the *Operands* of the language are referred to by using rendering artifacts (tables, row, columns and sheets).

A key requirement for the DPM-XL has been that the language that is currently used should change as little as possible. So, in practice, the formalization of the language has been a reverse-engineering process, where the starting point where the validations published by the EBA and EIOPA.

The main reasons for having **DPM-ML** are that it is providing a structured version of the calculations, with no need to parse expressions, and that it is referring to *Variables*, which are stable over time, because they represent business concepts, instead of rendering objects, which change often over time and do not hold any business meaning, but just a representation.

The DPM-ML is automatically derived from the DPM-XL and, eventually, can be also derived from languages other than the DPM-XL. DPM-ML relations to languages other than DPM-XL is out of the scope for this documentation. Instead, the documentation provides details on how *Operations* are represented both in DPM-XL and DPM-ML.
The DPM-XL and DPM-ML share the same *Operators* and the same information model.

## How

The formalization of the language has three pillars:

1. **Information model**: Specifies the artifacts that the language is using.
2. **Grammar**: Technical definition of the syntax of the language. Allows developers to build parsers for the language. It is provided in a separate file following the EBNF (Extended Backus-Naur form) notation (Only applicable to DPM-XL).
3. **Semantic specification**: Semantics for all *Operators* of the language, specifying formally their constraints and behaviour.

This document deals with the information model, while the grammar and semantic specification are provided as separate documents.

The Validation and Transformation Language specification has been used as an inspiration for this specification.
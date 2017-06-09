---
title: Operator Precedence
keywords: vblr6.chm1008988
f1_keywords:
- vblr6.chm1008988
ms.prod: office
ms.assetid: 05bd8527-1bf6-c2ab-6dac-f060d061cace
ms.date: 06/08/2017
---


# Operator Precedence



When several operations occur in an [expression](vbe-glossary.md), each part is evaluated and resolved in a predetermined order called operator precedence.
When expressions contain operators from more than one category, arithmetic operators are evaluated first, [comparison operators](vbe-glossary.md) are evaluated next, and logical operators are evaluated last. Comparison operators all have equal precedence; that is, they are evaluated in the left-to-right order in which they appear. Arithmetic and logical operators are evaluated in the following order of precedence:


|**Arithmetic**|**Comparison**|**Logical**|
|:-----|:-----|:-----|
|Exponentiation ( **^** )|Equality ( **=** )|**Not**|
|Negation ( **-** )|Inequality ( **&lt;&gt;** )|**And**|
|Multiplication and division ( *****, **/** )|Less than ( **&lt;** )|**Or**|
|Integer division ( **\** )|Greater than ( **&gt;** )|**Xor**|
|Modulus arithmetic ( **Mod** )|Less than or equal to ( **&lt;=** )|**Eqv**|
|Addition and subtraction ( **+**, **-** )|Greater than or equal to ( **>=** )|**Imp**|
|String concatenation ( **&;** )|**LikeIs**||
When multiplication and division occur together in an expression, each operation is evaluated as it occurs from left to right. When addition and subtraction occur together in an expression, each operation is evaluated in order of appearance from left to right. Parentheses can be used to override the order of precedence and force some parts of an expression to be evaluated before others. Operations within parentheses are always performed before those outside. Within parentheses, however, operator precedence is maintained.
The string concatenation operator ( **&;** ) is not an arithmetic operator, but in precedence, it does follow all arithmetic operators and precede all comparison operators.
The  **Like** operator is equal in precedence to all comparison operators, but is actually a pattern-matching operator.
The  **Is** operator is an object reference comparison operator. It does not compare objects or their values; it checks only to determine if two object references refer to the same object.


---
title: Mod Operator
keywords: vblr6.chm1008976
f1_keywords:
- vblr6.chm1008976
ms.prod: office
ms.assetid: cc1afd5d-ea12-a1df-3ffe-0d58f4d1e0ac
ms.date: 06/08/2017
---


# Mod Operator



Used to divide two numbers and return only the remainder.
 **Syntax**
 _result_**=**_number1_**Mod**_number2_
The  **Mod** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _number1_|Required; any [numeric expression](vbe-glossary.md).|
| _number2_|Required; any numeric expression.|
 **Remarks**
The modulus, or remainder, operator divides  _number1_ by _number2_ (rounding floating-point numbers to integers) and returns only the remainder as _result_. For example, in the following[expression](vbe-glossary.md), A ( _result_ ) equals 5.
Usually, the [data type](vbe-glossary.md) of _result_ is a[Byte](vbe-glossary.md),  **Byte** variant,[Integer](vbe-glossary.md),  **Integer** variant,[Long](vbe-glossary.md), or [Variant](vbe-glossary.md) containing a **Long**, regardless of whether or not _result_ is a whole number. Any fractional portion is truncated. However, if any expression is[Null](vbe-glossary.md),  _result_ is **Null**. Any expression that is[Empty](vbe-glossary.md) is treated as 0.

## Example

This example uses the  **Mod** operator to divide two numbers and return only the remainder. If either number is a floating-point number, it is first rounded to an integer.


```vb
Dim MyResult
MyResult = 10 Mod 5    ' Returns 0.
MyResult = 10 Mod 3    ' Returns 1.
MyResult = 12 Mod 4.3    ' Returns 0.
MyResult = 12.6 Mod 5    ' Returns 3.
```



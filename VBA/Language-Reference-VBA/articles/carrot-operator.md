---
title: ^ Operator
keywords: vblr6.chm1008849
f1_keywords:
- vblr6.chm1008849
ms.prod: office
ms.assetid: 9a2f874a-bf55-ae06-cf93-951d774eff0a
ms.date: 06/08/2017
---


# ^ Operator



Used to raise a number to the power of an exponent.
 **Syntax**
 _result_**=**_number_**^**_exponent_
The  **^** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _number_|Required; any [numeric expression](vbe-glossary.md).|
| _exponent_|Required; any numeric expression.|
 **Remarks**
A  _number_ can be negative only if _exponent_ is an integer value. When more than one exponentiation is performed in a single[expression](vbe-glossary.md), the  **^** operator is evaluated as it is encountered from left to right.
Usually, the [data type](vbe-glossary.md) of _result_ is a[Double](vbe-glossary.md) or a[Variant](vbe-glossary.md) containing a **Double**. However, if either _number_ or _exponent_ is a[Null](vbe-glossary.md) expression, _result_ is **Null**.

## Example

This example uses the  **^** operator to raise a number to the power of an exponent.


```vb
Dim MyValue
MyValue = 2 ^ 2    ' Returns 4.
MyValue = 3 ^ 3 ^ 3    ' Returns 19683.
MyValue = (-5) ^ 3    ' Returns -125.


```



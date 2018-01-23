---
title: "* Operator"
keywords: vblr6.chm1008844
f1_keywords:
- vblr6.chm1008844
ms.prod: office
ms.assetid: f45e939e-ff1d-c152-ad82-099e8f00ee8c
ms.date: 06/08/2017
---


# * Operator



Used to multiply two numbers.
 **Syntax**
 _result_**=**_number1_ * _number2_
The  **\*** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](vbe-glossary.md).|
| _number1_|Required; any [numeric expression](vbe-glossary.md).|
| _number2_|Required; any numeric expression.|
 **Remarks**
The [data type](vbe-glossary.md) of _result_ is usually the same as that of the most precise[expression](vbe-glossary.md). The order of precision, from least to most precise, is [Byte](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Single](vbe-glossary.md), [Currency](vbe-glossary.md), [Double](vbe-glossary.md), and [Decimal](vbe-glossary.md). The following are exceptions to this order:


|**If**|**Then  _result_ is**|
|:-----|:-----|
|Multiplication involves a  **Single** and a **Long**,|converted to a  **Double**.|
|The data type of  _result_ is a **Long**, **Single**, or **Date** variant that overflows its legal range,|converted to a  **Variant** containing a **Double**.|
|The data type of  _result_ is a **Byte** variant that overflows its legal range,|converted to an  **Integer** variant.|
|the data type of  _result_ is an **Integer** variant that overflows its legal range,|converted to a  **Long** variant.|
If one or both expressions are [Null](vbe-glossary.md) expressions, _result_ is **Null**. If an expression is[Empty](vbe-glossary.md), it is treated as 0.

 **Note**  The order of precision used by multiplication is not the same as the order of precision used by addition and subtraction.


## Example

This example uses the  **\*** operator to multiply two numbers.


```vb
Dim MyValue
MyValue = 2 * 2    ' Returns 4.
MyValue = 459.35 * 334.90     ' Returns 153836.315.


```



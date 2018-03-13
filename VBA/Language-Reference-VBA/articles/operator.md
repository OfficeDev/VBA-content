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
 <strong>Syntax</strong>
 
<em>result</em><strong>=</strong><em>number1</em> * <em>number2</em>
The  
<strong>\</strong>* operator syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                         |
|:----------------------|:-----------------------------------------------------|
| <em>result</em>       | Required; any numeric [variable](vbe-glossary.md).   |
| <em>number1</em>      | Required; any [numeric expression](vbe-glossary.md). |
| <em>number2</em>      | Required; any numeric expression.                    |

 **Remarks**
The [data type](vbe-glossary.md) of _result_ is usually the same as that of the most precise[expression](vbe-glossary.md). The order of precision, from least to most precise, is [Byte](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Single](vbe-glossary.md), [Currency](vbe-glossary.md), [Double](vbe-glossary.md), and [Decimal](vbe-glossary.md). The following are exceptions to this order:


| <strong>If</strong>                                                                                                                                     | <strong>Then  <em>result</em> is</strong>                                      |
|:--------------------------------------------------------------------------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------|
| Multiplication involves a  <strong>Single</strong> and a <strong>Long</strong>,                                                                         | converted to a  <strong>Double</strong>.                                       |
| The data type of  <em>result</em> is a <strong>Long</strong>, <strong>Single</strong>, or <strong>Date</strong> variant that overflows its legal range, | converted to a  <strong>Variant</strong> containing a <strong>Double</strong>. |
| The data type of  <em>result</em> is a <strong>Byte</strong> variant that overflows its legal range,                                                    | converted to an  <strong>Integer</strong> variant.                             |
| the data type of  <em>result</em> is an <strong>Integer</strong> variant that overflows its legal range,                                                | converted to a  <strong>Long</strong> variant.                                 |

If one or both expressions are [Null](vbe-glossary.md) expressions, _result_ is **Null**. If an expression is[Empty](vbe-glossary.md), it is treated as 0.

 **Note**  The order of precision used by multiplication is not the same as the order of precision used by addition and subtraction.


## Example

This example uses the  <strong>\</strong>* operator to multiply two numbers.


```vb
Dim MyValue
MyValue = 2 * 2    ' Returns 4.
MyValue = 459.35 * 334.90     ' Returns 153836.315.
```



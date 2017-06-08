---
title: IsDate Function
keywords: vblr6.chm1008951
f1_keywords:
- vblr6.chm1008951
ms.prod: office
ms.assetid: 832486a7-c69f-8d3b-f0fc-2f6a2f707ecc
ms.date: 06/08/2017
---


# IsDate Function



Returns  **True** if the expression is a date or is recognizable as a valid date or time; otherwise, it returns **False**.
 **Syntax**
 **IsDate(**_expression_**)**
The required  _expression_[argument](vbe-glossary.md) is a[Variant](vbe-glossary.md) containing a[date expression](vbe-glossary.md) or[string expression](vbe-glossary.md) recognizable as a date or time.
 **Remarks**
In Microsoft Windows, the range of valid dates is January 1, 100 A.D. through December 31, 9999 A.D.; the ranges vary among operating systems.

## Example

This example uses the  **IsDate** function to determine if an expression is recognized as a date or time value.


```vb
Dim MyVar, MyCheck
MyVar = "04/28/2014"    ' Assign valid date value.
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "April 28, 2014"    ' Assign valid date value.
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "13/32/2014"    ' Assign invalid date value.
MyCheck = IsDate(MyVar)    ' Returns False.

MyVar = "04.28.14"    ' Assign valid time value
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "04.28.2014"    ' Assign invalid time value.
MyCheck = IsDate(MyVar)    ' Returns False.

```



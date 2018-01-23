---
title: IsNumeric Function
keywords: vblr6.chm1008954
f1_keywords:
- vblr6.chm1008954
ms.prod: office
ms.assetid: b8184a41-8400-1228-c40f-1414eb4b6e63
ms.date: 06/08/2017
---


# IsNumeric Function



Returns a  **Boolean** value indicating whether an [expression](vbe-glossary.md) can be evaluated as a number.

 ## Syntax
 
 **IsNumeric(**_expression_**)**
 
The required  _expression_ [argument](vbe-glossary.md) is a [Variant](vbe-glossary.md) containing a [numeric expression](vbe-glossary.md) or [string expression](vbe-glossary.md).

 **Remarks**
 
 **IsNumeric** returns **True** if the entire _expression_ is recognized as a number; otherwise, it returns **False**.
 **IsNumeric** returns **False** if _expression_ is a [date expression](vbe-glossary.md).

## Example

This example uses the  **IsNumeric** function to determine if a variable can be evaluated as a number.


```vb
Dim MyVar, MyCheck
MyVar = "53"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns True.

MyVar = "459.95"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns True.

MyVar = "45 Help"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns False.


```



---
title: IsError Function
keywords: vblr6.chm1008824
f1_keywords:
- vblr6.chm1008824
ms.prod: office
ms.assetid: 7eab8dd7-6719-3fc1-fea2-3140cc6a0e5f
ms.date: 06/08/2017
---


# IsError Function



Returns a  **Boolean** value indicating whether an[expression](vbe-glossary.md) is an error value.
 **Syntax**
 **IsError(**_expression_**)**
The required  _expression_[argument](vbe-glossary.md) can be any valid expression.
 **Remarks**
Error values are created by converting real numbers to error values using the  **CVErr** function. The **IsError** function is used to determine if a[numeric expression](vbe-glossary.md) represents an error. **IsError** returns **True** if the _expression_ argument indicates an error; otherwise, it returns **False**.

## Example

This example uses the  **IsError** function to check if a numeric expression is an error value. The **CVErr** function is used to return an **Error Variant** from a user-defined function. Assume `UserFunction` is a user-defined function procedure that returns an error value; for example, a return value assigned with the statement `UserFunction = CVErr(32767)`, where 32767 is a user-defined number.


```vb
Dim ReturnVal, MyCheck
ReturnVal = UserFunction()
MyCheck = IsError(ReturnVal)    ' Returns True.
```



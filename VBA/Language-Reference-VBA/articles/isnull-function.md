---
title: IsNull Function
keywords: vblr6.chm1008953
f1_keywords:
- vblr6.chm1008953
ms.prod: office
ms.assetid: 875909ba-289e-aba9-0462-9327efe0bc46
ms.date: 06/08/2017
---


# IsNull Function



Returns a  **Boolean** value that indicates whether an[expression](vbe-glossary.md) contains no valid data ([Null](vbe-glossary.md)).
 **Syntax**
 **IsNull(**_expression_**)**
The required  _expression_[argument](vbe-glossary.md) is a[Variant](vbe-glossary.md) containing a[numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md).
 **Remarks**
 **IsNull** returns **True** if _expression_ is **Null**; otherwise, **IsNull** returns **False**. If _expression_ consists of more than one[variable](vbe-glossary.md),  **Null** in any constituent variable causes **True** to be returned for the entire expression.
The  **Null** value indicates that the **Variant** contains no valid data. **Null** is not the same as[Empty](vbe-glossary.md), which indicates that a variable has not yet been initialized. It is also not the same as a zero-length string (""), which is sometimes referred to as a null string.


 **Important**  Use the  **IsNull** function to determine whether an expression contains a **Null** value. Expressions that you might expect to evaluate to **True** under some circumstances, such as `If Var = Null` and `If Var <> Null`, are always  **False**. This is because any expression containing a **Null** is itself **Null** and, therefore, **False**.



## Example

This example uses the  **IsNull** function to determine if a variable contains a **Null**.


```vb
Dim MyVar, MyCheck
MyCheck = IsNull(MyVar)    ' Returns False.

MyVar = ""
MyCheck = IsNull(MyVar)    ' Returns False.

MyVar = Null
MyCheck = IsNull(MyVar)    ' Returns True.


```



---
title: IsObject Function
keywords: vblr6.chm1008825
f1_keywords:
- vblr6.chm1008825
ms.prod: office
ms.assetid: 24fee32f-52ed-48b3-a52e-9a66b0e62723
ms.date: 06/08/2017
---


# IsObject Function



Returns a  **Boolean** value indicating whether an [identifier](vbe-glossary.md) represents an object [variable](vbe-glossary.md).
 **Syntax**
 **IsObject(**_identifier_**)**
The required  _identifier_ [argument](vbe-glossary.md) is a variable name.
 **Remarks**
 **IsObject** is useful only in determining whether a [Variant](vbe-glossary.md) is of **VarType** **vbObject**. This could occur if the **Variant** actually references (or once referenced) an object, or if it contains **Nothing.**
 **IsObject** returns **True** if _identifier_ is a variable declared with [Object](vbe-glossary.md) type or any valid [class](vbe-glossary.md) type, or if _identifier_ is a **Variant** of **VarType** **vbObject**, or a user-defined object; otherwise, it returns **False**. **IsObject** returns **True** even if the variable has been set to **Nothing**.
Use error trapping to be sure that an object reference is valid.

## Example

This example uses the  **IsObject** function to determine if an identifier represents an object variable. `MyObject` and and `YourObject` are object variables of the same type. They are generic names used for illustration purposes only.


```vb
Dim MyInt As Integer              ' Declare variables.
Dim YourObject, MyCheck           ' Note: default variable type is Variant
Dim MyObject As Object
Set YourObject = MyObject         ' Assign an object reference.
MyCheck = IsObject(YourObject)    ' Returns True.
MyCheck = IsObject(MyInt)         ' Returns False.
MyCheck = IsObject(Nothing)       ' Returns True.
MyCheck = IsObject(Empty)         ' Returns False.
MyCheck = IsObject(Null)          ' Returns False.
```



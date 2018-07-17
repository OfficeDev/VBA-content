---
title: Can't assign or coerce array of fixed-length string or user-defined type to Variant
keywords: vblr6.chm1040123
f1_keywords:
- vblr6.chm1040123
ms.prod: office
ms.assetid: 4b33f923-6e76-491c-2b0a-901045d03ff8
ms.date: 06/08/2017
---


# Can't assign or coerce array of fixed-length string or user-defined type to Variant

A  **Variant** can only accept assignment of data having a valid **VarType**. This error has the following causes and solutions:



- You tried to pass an [array](vbe-glossary.md) of fixed-length strings. When a single fixed-length string is assigned to a **Variant**, it's coerced to a variable-length string, but this can't be done for an array of fixed-length strings.
    
    If you must pass the array, use a loop to assign the individual elements of the array to the elements of a temporary array of variable-length strings. You can then assign the array to a variant and use  **Erase** to deallocate the temporary array. However, you can't deallocate a fixed-size array with **Erase**.
    
- You tried to pass a fixed-length string or [user-defined type](vbe-glossary.md) to the **VarType** function or **TypeName** function.
    
    An [argument](vbe-glossary.md) to the **VarType** or **TypeName** function must be a valid **Variant** type.
    
- You tried to assign a user-defined type to a  **Variant** variable. Although you can't directly assign a whole[variable](vbe-glossary.md) of user-defined type to a **Variant**, you can use the **Array** function to assign the individual elements of a variable of user-defined type to a **Variant**. This produces a **Variant** containing an array of variants. The **VarType** of each element in this array of variants corresponds to the original type of each element of the user-defined type.
    
- You tried to pass an [array](vbe-glossary.md) of fixed-length strings or user-defined types as an argument in a procedure call that requires a **Variant** argument. Note that any time a procedure is late bound, that is, when the call must be constructed at[run time](vbe-glossary.md), all arguments must be passed as  **Variant** types. For example, the following code causes this error:
    
```vb
Dim MyForm As Object    ' Because MyForm is Object, binding is late. 
Set MyForm = New Form1 
Dim StringArray(10) As String * 12 
' The next line generates the error. 
MyForm.MyProc StringArray 

  ```


    For the string array, use a loop to assign each individual member of the array to a temporary array of variable-length strings. You can then assign that array to a  **Variant** to pass to the procedure. For an array of user-defined types, you can use the **Array** function to assign the individual elements of a variable of user-defined type to a **Variant**. This produces a **Variant** containing an array of variants. The **VarType** of each element in this array of variants corresponds to the original type of each element of the user-defined type.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


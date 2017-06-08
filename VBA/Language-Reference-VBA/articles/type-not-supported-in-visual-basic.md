---
title: Type not supported in Visual Basic
keywords: vblr6.chm1015664
f1_keywords:
- vblr6.chm1015664
ms.prod: office
ms.assetid: 59add4f9-d15a-7342-e7fc-4b21420a5e41
ms.date: 06/08/2017
---


# Type not supported in Visual Basic

Not all types are supported in Visual Basic. This error has the following cause and solution:



- You tried to use a type in your program that has no equivalent in Visual Basic for Applications. For example, Visual Basic has no pointer or unsigned integer type, so if you try to create a [variable](vbe-glossary.md) of one of those types from an[object library](vbe-glossary.md), this error occurs. In the following example that follows, even though  `Rainbow` may be a valid structure, Visual Basic can't create a variable of that type if it contains a type Visual Basic doesn't recognize:
    
```vb
Dim MyVar As Rainbow    ' Causes error. 

  ```


    If the type is a valid [parameter](vbe-glossary.md) type for a function in an object library, this error means only that you can't create a variable of that type in your own code. Although you can't always declare variables with a[data type](vbe-glossary.md) specified in an object's documentation, there is often a Visual Basic equivalent. For example, although Visual Basic has no pointer type, you can pass a pointer to a function to an API function by using the **AddressOf** operator. Also, check the **Variant** type's subtypes. You can often use them as equivalents of types not offered directly in Visual Basic. In some cases, however, Visual Basic simply has no equivalent. For example, data pointers aren't available.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


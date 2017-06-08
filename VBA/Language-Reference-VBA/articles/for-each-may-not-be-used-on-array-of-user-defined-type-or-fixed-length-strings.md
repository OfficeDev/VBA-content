---
title: For Each may not be used on array of user-defined type or fixed-length strings
keywords: vblr6.chm1040138
f1_keywords:
- vblr6.chm1040138
ms.prod: office
ms.assetid: 37976c99-e8a7-250b-5b63-5d0fd204d576
ms.date: 06/08/2017
---


# For Each may not be used on array of user-defined type or fixed-length strings

 **For Each** constructs are only valid for[collections](vbe-glossary.md) and[arrays](vbe-glossary.md) of intrinsic types, including arrays of objects. Also, arrays of fixed-length strings can't be iterated using **For** **Each**. This error has the following causes and solutions:



- The elements of the array in your  **For Each** construct have a[user-defined type](vbe-glossary.md).
    
    Use an ordinary  **For...Next** loop to iterate the elements of the array.
    
- The elements of the array in your  **For Each** construct have a fixed-length string type. Use an ordinary **For...Next** loop to iterate the elements of the array.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


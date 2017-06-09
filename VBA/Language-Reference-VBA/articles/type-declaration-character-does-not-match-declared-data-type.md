---
title: Type-declaration character does not match declared data type
keywords: vblr6.chm1040102
f1_keywords:
- vblr6.chm1040102
ms.prod: office
ms.assetid: d3581bff-e345-a1ac-e092-7ccb993be618
ms.date: 06/08/2017
---


# Type-declaration character does not match declared data type

The [data type](vbe-glossary.md) of a[variable](vbe-glossary.md) can't be changed by appending the[type-declaration character](vbe-glossary.md) for another type. This error has the following cause and solution:



- You declared a variable of a specific type, referenced a variable of the same name in the same [scope](vbe-glossary.md), and then appended an inconsistent type-declaration character.
    
    If you want to be able to change the type of data assigned to a variable, declare the variable as a  **Variant**. If you simply appended an incorrect type-declaration character, delete or change it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


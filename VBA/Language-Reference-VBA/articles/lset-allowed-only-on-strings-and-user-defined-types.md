---
title: LSet allowed only on strings and user-defined types
keywords: vblr6.chm1011214
f1_keywords:
- vblr6.chm1011214
ms.prod: office
ms.assetid: 6e641999-66f1-46fb-869f-369d3f5274b8
ms.date: 06/08/2017
---


# LSet allowed only on strings and user-defined types

 **LSet** is used to left align data within strings and[variables](vbe-glossary.md) of[user-defined type](vbe-glossary.md). This error has the following causes and solutions:



- The specified variable isn't a string or user-defined type. If you are trying to block assign one array to another,  **LSet** does not work. You must use a loop to assign each element individually.
    
- You tried to use  **LSet** with an object. **LSet** can also be used to assign the elements of a user-defined type variable to the elements of a different, but compatible, user-defined type. Although objects are similar to user-defined types, you can't use **LSet** on them. Similarly, you can't use **LSet** on variables of user-defined types that contain strings, objects, or variants.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


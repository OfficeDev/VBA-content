---
title: RSet allowed only on strings
keywords: vblr6.chm1040145
f1_keywords:
- vblr6.chm1040145
ms.prod: office
ms.assetid: cf7a404b-de1f-501b-c961-011c46e460c8
ms.date: 06/08/2017
---


# RSet allowed only on strings

 **RSet** is used to right align string data within fixed-length or variable-length strings. This error has the following cause and solution:



- You tried to use the  **RSet** statement on a[variable](vbe-glossary.md) that isn't a string.
    
    If appropriate, try converting the variable to a string. Otherwise, don't use  **RSet**.
    
     **Note**  Although the  **LSet** statement can be used to assign the elements of one[user-defined type](vbe-glossary.md) variable to the elements of a different, but compatible, user-defined type, such assignments are discouraged because they can't be guaranteed to be portable.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


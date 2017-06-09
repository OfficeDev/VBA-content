---
title: Bad record number (Error 63)
keywords: vblr6.chm1011088
f1_keywords:
- vblr6.chm1011088
ms.prod: office
ms.assetid: 7535b68a-cb1f-a443-ab6c-640673de281d
ms.date: 06/08/2017
---


# Bad record number (Error 63)

An error occurred during the attempted file access. This error has the following cause and solution:



- The record number in a  **Put** or **Get** statement is less than or equal to zero. Check the calculations used in generating the record number. Make sure that the[variables](vbe-glossary.md) containing the record number or used in calculating record numbers are spelled correctly. A misspelled variable name is implicitly declared and initialized to zero, unless you have properly placed **Option Explicit** in the[module](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


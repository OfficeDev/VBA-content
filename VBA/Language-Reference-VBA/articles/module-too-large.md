---
title: Module too large
keywords: vblr6.chm1057026
f1_keywords:
- vblr6.chm1057026
ms.prod: office
ms.assetid: b00483e1-d3b2-f532-eaa3-fae61f45c013
ms.date: 06/08/2017
---


# Module too large

A [module](vbe-glossary.md) contains code within the[project](vbe-glossary.md). This error has the following cause and solution:



- There is too much code in the module.
    
    Create a new module and move some of the [procedures](vbe-glossary.md) from this module to the new one. If the current module contains[module-level](vbe-glossary.md) declarations of data that must be visible to the procedures in the new module, declare that data as **Public**.
    
     **Note**  [Comments](vbe-glossary.md) aren't counted as lines of code. Therefore, deleting comments doesn't prevent this error.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


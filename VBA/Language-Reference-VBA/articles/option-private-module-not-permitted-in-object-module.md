---
title: Option Private Module not permitted in object module
keywords: vblr6.chm1057033
f1_keywords:
- vblr6.chm1057033
ms.prod: office
ms.assetid: 4b3098a1-5bbd-61bf-f242-8b4e1b1714a2
ms.date: 06/08/2017
---


# Option Private Module not permitted in object module

 **Option Private Module** makes the contents of a[module](vbe-glossary.md) unavailable to other[projects](vbe-glossary.md), while preserving their availability to your project. This error has the following cause and solution:



- The statement  **Option Private Module** appears in an[object module](vbe-glossary.md).
    
    Remove the  **Option Private Module** statement from the module. Object modules have the characteristic of **Option Private Module** by default. Changing the default can't be done from code. See your[host application's](vbe-glossary.md) documentation for information on giving object module members wider visibility.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


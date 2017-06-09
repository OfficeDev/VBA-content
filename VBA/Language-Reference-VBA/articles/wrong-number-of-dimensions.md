---
title: Wrong number of dimensions
keywords: vblr6.chm1011078
f1_keywords:
- vblr6.chm1011078
ms.prod: office
ms.assetid: ccd07473-8199-d616-911d-3c16b2ffe218
ms.date: 06/08/2017
---


# Wrong number of dimensions

You must reference an [array](vbe-glossary.md) with indexes corresponding to the same number of dimensions as appear in the array's[declaration](vbe-glossary.md). This error has the following cause and solution:



- You referred to an array with a different number of dimensions than it actually contains. For example, referring to an array as  `X(2,4)` (an array with two dimensions) when it has been defined as `Dim X(5)` (an array with one), generates this error. Check the declaration of the array and, in the reference, include one index for each dimension in the declaration.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


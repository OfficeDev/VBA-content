---
title: For control variable already in use
keywords: vblr6.chm1011174
f1_keywords:
- vblr6.chm1011174
ms.prod: office
ms.assetid: 9b817917-5156-7dc6-f4f1-4fc6626ad5c9
ms.date: 06/08/2017
---


# For control variable already in use

When you nest  **For...Next** loops, you must use different control[variables](vbe-glossary.md) in each one. This error has the following cause and solution:



- An inner  **For** loop uses the same counter as an enclosing **For** loop. Check nested loops for repetition. For example, if the outer loop uses `For Count = 1 To 25`, the inner loops can't use  `Count` as the control variables.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


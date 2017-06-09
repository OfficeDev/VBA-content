---
title: Project contains too many procedure, variable, and constant names
keywords: vblr6.chm1018968
f1_keywords:
- vblr6.chm1018968
ms.prod: office
ms.assetid: d78ca072-6a1f-370a-2611-3f088b320a5a
ms.date: 06/08/2017
---


# Project contains too many procedure, variable, and constant names

A project's [procedure](vbe-glossary.md), [variable](vbe-glossary.md), [constant](vbe-glossary.md), and [parameter](vbe-glossary.md) names are stored in a name table. This error has the following cause and solution:



- The number of names in the project's name table exceeds 32,768. The name table may contain some temporary duplicates. You can compact the name table by saving the [project](vbe-glossary.md) to a disk, and then closing it. If the problem persists after you reopen the project, reduce the number of names by reusing local variable names in multiple procedures, and then recompact the table by saving the project, closing it, and reopening it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


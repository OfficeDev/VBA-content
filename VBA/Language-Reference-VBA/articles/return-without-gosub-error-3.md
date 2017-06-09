---
title: Return without GoSub (Error 3)
keywords: vblr6.chm1011266
f1_keywords:
- vblr6.chm1011266
ms.prod: office
ms.assetid: 396d3d0f-6af2-4709-bf3c-3a35668398d7
ms.date: 06/08/2017
---


# Return without GoSub (Error 3)

A  **Return** statement must have a corresponding **GoSub** statement. This error has the following cause and solution:



- You have a  **Return** statement that can't be matched with a **GoSub** statement. Make sure your **GoSub** statement wasn't inadvertently deleted.
    

Unlike  **For...Next**, **While...Wend**, and **Sub...End Sub**, which are matched at[compile time](vbe-glossary.md),  **GoSub** and **Return** are matched at[run time](vbe-glossary.md).
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


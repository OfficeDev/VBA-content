---
title: Case without Select Case
keywords: vblr6.chm1011109
f1_keywords:
- vblr6.chm1011109
ms.prod: office
ms.assetid: 59d1eb92-b346-4013-b4fa-e99e40c2a9d6
ms.date: 06/08/2017
---


# Case without Select Case

A  **Case** statement must occur within a **Select Case...End Select Block**. This error has the following cause and solution:



- A  **Case** statement can't be matched with a preceding **Select Case** statement. Check other control structures within the **Select Case...Case** structure and verify that they are correctly matched. For example, an **If** without a matching **End If** inside the **Select Case...End Select** structure generates this error.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


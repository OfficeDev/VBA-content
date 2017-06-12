---
title: Must be first statement on the line
keywords: vblr6.chm1040049
f1_keywords:
- vblr6.chm1040049
ms.prod: office
ms.assetid: 5aa6b5a6-27ed-7825-f204-20b9697f25f3
ms.date: 06/08/2017
---


# Must be first statement on the line

Not all [keywords](vbe-glossary.md) can appear at the beginning of a line of code. This error has the following causes and solutions:



- You preceded a  **Sub**, **Function**, or **Property** statement with another statement on the same line. A **Sub**, **Function**, or **Property** statement must always be the first statement on any line in which it appears (unless preceded by the keyword **Public**, **Private**, or **Static** ).
    
- You preceded an  **End If**, **Else**, or **ElseIf** statement with another statement on the same line. An **End If**, **Else**, or **ElseIf** (only when used in a block **If** structure) statement must always be the first statement on any line in which it appears.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


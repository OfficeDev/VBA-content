---
title: Only comments may appear after End Sub, End Function, or End Property
keywords: vblr6.chm1040081
f1_keywords:
- vblr6.chm1040081
ms.prod: office
ms.assetid: 6268c6e6-1bd6-d7f8-50e3-a749bb578bcf
ms.date: 06/08/2017
---


# Only comments may appear after End Sub, End Function, or End Property

Only [comments](vbe-glossary.md), directives, and [declarations](vbe-glossary.md) are permitted outside[procedures](vbe-glossary.md). This error has the following cause and solution:



- You placed executable code outside a procedure. Any nondeclarative lines outside a procedure must begin with a comment delimiter ( **'** ). Declarative statements must appear before the first procedure declaration. Comments are ignored when the code executes.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


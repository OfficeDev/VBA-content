---
title: Selected watch expression invalid
keywords: vblr6.chm1040216
f1_keywords:
- vblr6.chm1040216
ms.prod: office
ms.assetid: b5a05c94-4ec4-f92b-6073-1635ed49ca69
ms.date: 06/08/2017
---


# Selected watch expression invalid

It isn't always possible to select a valid [watch expression](vbe-glossary.md). This error has the following causes and solutions:



- You chose the  **Instant Watch** command, but the selected[expression](vbe-glossary.md) isn't a valid expression. For example, you can't watch a[comment](vbe-glossary.md) or a **Sub** procedure call.
    
    Select the expression in such a way that it is valid, or choose  **Add Watch** and type in a valid expression.
    
- The watch expression must have code syntax corresponding to the [locale](vbe-glossary.md) of the[project](vbe-glossary.md) that defines the expression being watched.
    
    Rewrite the expression in a way that is valid for the locale.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


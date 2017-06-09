---
title: Expression too complex (Error 16)
keywords: vblr6.chm1076493
f1_keywords:
- vblr6.chm1076493
ms.prod: office
ms.assetid: 718b5c52-5844-fa60-4490-6db2529dcc4e
ms.date: 06/08/2017
---


# Expression too complex (Error 16)

The number of subexpressions allowed in a floating-point [expression](vbe-glossary.md) varies among platforms. For example, on 32-bit Microsoft Windows operating systems, the limit is 8 levels of nested floating-point expressions. This error has the following cause and solution:



- A floating-point expression contains too many nested subexpressions.
    
    Break the expression into as many separate expressions as necessary to prevent the error from occurring.
    
     **Note**  In earlier versions of Visual Basic, Error 16 was "String expression too complex." That error condition can no longer occur. However, if you have early code that traps and handles that error, you should remove it to prevent confusion with this new error.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


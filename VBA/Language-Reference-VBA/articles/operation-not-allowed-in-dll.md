---
title: Operation not allowed in DLL
keywords: vblr6.chm1111923
f1_keywords:
- vblr6.chm1111923
ms.prod: office
ms.assetid: ff4949cc-44ff-085c-3343-9b9a1ee8e2ad
ms.date: 06/08/2017
---


# Operation not allowed in DLL

Not all Visual Basic statements are allowed within a [dynamic-link library (DLL)](vbe-glossary.md). This error has the following causes and solutions:



- You tried to create a DLL from a [class](vbe-glossary.md) that contains a statement that can't be executed from a DLL.
    
    Check your class and remove any statements that can't be executed within a DLL, for example, the  **End statement**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


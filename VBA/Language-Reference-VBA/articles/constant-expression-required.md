---
title: Constant expression required
keywords: vblr6.chm1040114
f1_keywords:
- vblr6.chm1040114
ms.prod: office
ms.assetid: e0493fe4-8f50-c935-391f-0ffaca726b2b
ms.date: 06/08/2017
---


# Constant expression required

A [constant](vbe-glossary.md) must be initialized. This error has the following causes and solutions:



- You tried to initialize a constant with a [variable](vbe-glossary.md), an instance of a [user-defined type](vbe-glossary.md), an object, or the return value of a function call.
    
    Initialize constants with literals, previously declared constants, or literals and constants joined by operators (except the  **Is** logical operator).
    
- [array](vbe-glossary.md)
    
    To declare a dynamic array within a [procedure](vbe-glossary.md), declare the array with  **ReDim** and specify the number of elements with a variable.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


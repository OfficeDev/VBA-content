---
title: Object doesn't support this property or method (Error 438)
keywords: vblr6.chm1011328
f1_keywords:
- vblr6.chm1011328
ms.prod: office
ms.assetid: 0fbab746-dc6d-b227-429a-1f56bb4ca448
ms.date: 06/08/2017
---


# Object doesn't support this property or method (Error 438)

Not all objects support all [properties](vbe-glossary.md) and[methods](vbe-glossary.md). This error has the following cause and solution:



- You specified a method or property that doesn't exist for this [Automation object](vbe-glossary.md).
    
    See the object's documentation for more information on the object and check the spellings of properties and methods.
    
- You specified a  **Friend** procedure to be called late bound. The name of a **Friend** procedure must be known at[compile time](vbe-glossary.md). It can't appear in a late-bound call.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


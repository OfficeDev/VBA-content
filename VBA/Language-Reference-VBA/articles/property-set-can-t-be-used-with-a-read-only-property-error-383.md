---
title: Property Set can't be used with a read-only property (Error 383)
keywords: vblr6.chm1000383
f1_keywords:
- vblr6.chm1000383
ms.prod: office
ms.assetid: 42ea9723-86e1-7409-844e-9bda4be80c5f
ms.date: 06/08/2017
---


# Property Set can't be used with a read-only property (Error 383)

It may not be possible to obtain a reference to a [property](vbe-glossary.md) at[run time](vbe-glossary.md).This error has the following cause and solution:



- You tried to get a reference to a property that's read-only at run time. Since you can use a reference for both reading and writing, the property must provide run-time support for both operations for a reference to be obtained at run time. You can only use a  **Property Get** with this property.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


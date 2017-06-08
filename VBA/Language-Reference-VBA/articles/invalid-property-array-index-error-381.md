---
title: Invalid property-array index (Error 381)
keywords: vblr6.chm1107958
f1_keywords:
- vblr6.chm1107958
ms.prod: office
ms.assetid: 63598821-9427-e71d-2168-d4448a684005
ms.date: 06/08/2017
---


# Invalid property-array index (Error 381)

A [property](vbe-glossary.md) value may consist of an[array](vbe-glossary.md) of values. This error has the following cause and solution:



- A component's property array could have a lower bound of zero and an upper bound equal to the number of elements in the array minus 1. Alternatively, the lower bound could be 1 and the upper bound could equal the number of elements in the array. Check the component's documentation to make sure your index is within the valid range for the specified property.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


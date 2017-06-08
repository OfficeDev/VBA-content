---
title: Argument required for Property Let or Property Set
keywords: vblr6.chm1040125
f1_keywords:
- vblr6.chm1040125
ms.prod: office
ms.assetid: bbefad41-c17c-d1d3-52ac-32389acb3b7f
ms.date: 06/08/2017
---


# Argument required for Property Let or Property Set

The purpose of  **Property Let** and **Property Set** procedures is to give a new value to a[property](vbe-glossary.md). This error has the following causes and solutions:



- In setting the property, the value doesn't appear in the right place. Place the value to which you want to set the property on the right side of the [expression](vbe-glossary.md) setting the property value.
    
- In the procedure definition, the [parameter](vbe-glossary.md) defined to receive the value passed on the right side of the expression is missing or misplaced.
    
    Specify a parameter for the value argument list in the procedure definition. If the procedure takes more than one [argument](vbe-glossary.md), the property-value parameter must appear last in the list.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


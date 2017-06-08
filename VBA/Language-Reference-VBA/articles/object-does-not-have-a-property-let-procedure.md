---
title: Object does not have a Property Let procedure
keywords: vblr6.chm1115313
f1_keywords:
- vblr6.chm1115313
ms.prod: office
ms.assetid: 88b10e14-bd01-3738-2509-f98dff5dd0e7
ms.date: 06/08/2017
---


# Object does not have a Property Let procedure

You can't assign a value to a [property](vbe-glossary.md) unless it has exposed a **Property Let** method. This error has the following causes and solutions:



- You tried to assign a value to a property that hasn't exposed a  **Property Let** method. You can't directly assign a value to this property. If you created the[class](vbe-glossary.md), you can modify the interface by exposing a  **Property Let** method. Otherwise, check the component's documentation to determine if there is an indirect method for assigning the value.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


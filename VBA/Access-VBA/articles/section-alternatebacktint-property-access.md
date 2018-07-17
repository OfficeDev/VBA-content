---
title: Section.AlternateBackTint Property (Access)
keywords: vbaac10.chm14608
f1_keywords:
- vbaac10.chm14608
ms.prod: access
api_name:
- Access.Section.AlternateBackTint
ms.assetid: 7758713d-cfba-ac57-91c7-fcdab26ae44a
ms.date: 06/08/2017
---


# Section.AlternateBackTint Property (Access)

Gets or sets the tint applied to the theme color in the  **AlternateBackColor** property of the section. Read/write **Single**.


## Syntax

 _expression_. **AlternateBackTint**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **AlternateBackTint** property contains a numeric expression that can be used to lighten the theme color in the **AlternateBackColor** property. The default value of the **AlternateBackTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[Section Object](section-object-access.md)


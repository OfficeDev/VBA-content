---
title: ObjectFrame.GridlineTint Property (Access)
keywords: vbaac10.chm14636
f1_keywords:
- vbaac10.chm14636
ms.prod: access
api_name:
- Access.ObjectFrame.GridlineTint
ms.assetid: f32466e0-0924-97c7-2454-7632730ffcfa
ms.date: 06/08/2017
---


# ObjectFrame.GridlineTint Property (Access)

Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **GridlineTint**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **GridlineTint** property contains a numeric expression that can be used to lighten the theme color in the **GridlineColor** property. The default value of the **GridlineTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)


---
title: CommandButton.PressedTint Property (Access)
keywords: vbaac10.chm14621
f1_keywords:
- vbaac10.chm14621
ms.prod: access
api_name:
- Access.CommandButton.PressedTint
ms.assetid: 11439c75-f951-a551-12ee-b7b2d2e8ee94
ms.date: 06/08/2017
---


# CommandButton.PressedTint Property (Access)

Gets or sets the tint applied to the theme color in the  **PressedColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **PressedTint**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **PressedTint** property contains a numeric expression that can be used to lighten the theme color in the **PressedColor** property. The default value of the **PressedTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)


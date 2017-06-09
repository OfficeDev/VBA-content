---
title: TabControl.PressedTint Property (Access)
keywords: vbaac10.chm14621
f1_keywords:
- vbaac10.chm14621
ms.prod: access
api_name:
- Access.TabControl.PressedTint
ms.assetid: 1826cb99-d49c-465c-6c80-bca5a31f0f06
ms.date: 06/08/2017
---


# TabControl.PressedTint Property (Access)

Gets or sets the tint applied to the theme color in the  **PressedColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **PressedTint**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **PressedTint** property contains a numeric expression that can be used to lighten the theme color in the **PressedColor** property. The default value of the **PressedTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)


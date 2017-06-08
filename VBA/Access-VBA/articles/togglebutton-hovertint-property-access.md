---
title: ToggleButton.HoverTint Property (Access)
keywords: vbaac10.chm14613
f1_keywords:
- vbaac10.chm14613
ms.prod: access
api_name:
- Access.ToggleButton.HoverTint
ms.assetid: fbdb27bb-8a21-729c-17d6-a0e9b43826ae
ms.date: 06/08/2017
---


# ToggleButton.HoverTint Property (Access)

Gets or sets the tint applied to the theme color in the  **HoverColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **HoverTint**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **HoverTint** property contains a numeric expression that can be used to lighten the theme color in the **HoverColor** property. The default value of the **HoverTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)


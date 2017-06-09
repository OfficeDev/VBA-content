---
title: ToggleButton.PressedShade Property (Access)
keywords: vbaac10.chm14622
f1_keywords:
- vbaac10.chm14622
ms.prod: access
api_name:
- Access.ToggleButton.PressedShade
ms.assetid: 72176e9c-68bf-971c-3147-fea692240d17
ms.date: 06/08/2017
---


# ToggleButton.PressedShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **PressedColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **PressedShade**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **PressedShade** property contains a numeric expression that can be used to darken the theme color in the **PressedColor** property. The default value of the **PressedShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and store the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and store the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)


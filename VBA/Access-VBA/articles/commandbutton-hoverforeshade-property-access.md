---
title: CommandButton.HoverForeShade Property (Access)
keywords: vbaac10.chm14618
f1_keywords:
- vbaac10.chm14618
ms.prod: access
api_name:
- Access.CommandButton.HoverForeShade
ms.assetid: be9e6008-4cc4-94b5-869e-068c3b73443a
ms.date: 06/08/2017
---


# CommandButton.HoverForeShade Property (Access)

Gets or sets the shade applied to the theme color in the  **HoverForeColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **HoverForeShade**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **HoverForeShade** property contains a numeric expression that can be used to darken the theme color in the **HoverForeColor** property. The default value of the **HoverForeShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)


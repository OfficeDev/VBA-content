---
title: CommandButton.HoverShade Property (Access)
keywords: vbaac10.chm14614
f1_keywords:
- vbaac10.chm14614
ms.prod: access
api_name:
- Access.CommandButton.HoverShade
ms.assetid: 9a8b86d0-3849-9902-4dbf-c911c7fbe8e2
ms.date: 06/08/2017
---


# CommandButton.HoverShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **HoverColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **HoverShade**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **HoverShade** property contains a numeric expression that can be used to darken the theme color in the **HoverColor** property. The default value of the **HoverShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and store the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and store the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)


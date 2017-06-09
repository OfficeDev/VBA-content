---
title: TextBox.GridlineShade Property (Access)
keywords: vbaac10.chm14637
f1_keywords:
- vbaac10.chm14637
ms.prod: access
api_name:
- Access.TextBox.GridlineShade
ms.assetid: 33daf4ec-1587-63c8-4b23-2abdf5087bbe
ms.date: 06/08/2017
---


# TextBox.GridlineShade Property (Access)

Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **GridlineShade**

 _expression_ A variable that represents an **TextBox** object.


## Remarks

The  **GridlineShade** property contains a numeric expression that can be used to darken the theme color in the **GridlineColor** property. The default value of the **GridlineShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and store the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and store the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)


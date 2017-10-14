---
title: ListBox.GridlineShade Property (Access)
keywords: vbaac10.chm14637
f1_keywords:
- vbaac10.chm14637
ms.prod: access
api_name:
- Access.ListBox.GridlineShade
ms.assetid: 261c89f0-e1d8-41a0-6d8b-97332a6a01db
ms.date: 06/08/2017
---


# ListBox.GridlineShade Property (Access)

Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **GridlineShade**

 _expression_ A variable that represents an **ListBox** object.


## Remarks

The  **GridlineShade** property contains a numeric expression that can be used to darken the theme color in the **GridlineColor** property. The default value of the **GridlineShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and store the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and store the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)


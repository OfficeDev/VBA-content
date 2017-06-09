---
title: TextBox.BackShade Property (Access)
keywords: vbaac10.chm14633
f1_keywords:
- vbaac10.chm14633
ms.prod: access
api_name:
- Access.TextBox.BackShade
ms.assetid: 36db2540-6d5b-ed43-a303-70b6282398cf
ms.date: 06/08/2017
---


# TextBox.BackShade Property (Access)

Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BackShade**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **BackShade** property contains a numeric expression that can be used to darken the theme color in the **BackColor** property. The default value of the **BackShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens  **BackColor** by 75%.


```vb
Me.ctl.BackShade=25
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)


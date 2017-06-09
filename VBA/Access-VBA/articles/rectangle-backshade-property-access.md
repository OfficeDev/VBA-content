---
title: Rectangle.BackShade Property (Access)
keywords: vbaac10.chm14633
f1_keywords:
- vbaac10.chm14633
ms.prod: access
api_name:
- Access.Rectangle.BackShade
ms.assetid: 917bbe83-940a-edc5-8c6d-230af220d641
ms.date: 06/08/2017
---


# Rectangle.BackShade Property (Access)

Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BackShade**

 _expression_ A variable that represents a **Rectangle** object.


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


[Rectangle Object](rectangle-object-access.md)


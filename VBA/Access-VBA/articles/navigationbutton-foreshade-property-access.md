---
title: NavigationButton.ForeShade Property (Access)
keywords: vbaac10.chm14606
f1_keywords:
- vbaac10.chm14606
ms.prod: access
api_name:
- Access.NavigationButton.ForeShade
ms.assetid: 312eb5f5-0ed8-e3da-0720-420ed03ee19c
ms.date: 06/08/2017
---


# NavigationButton.ForeShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **ForeShade**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

The  **ForeShade** property contains a numeric expression that can be used to darken the theme color in the **BackColor** property. The default value of the **ForeShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens the  **ForeColor** property by 75%.


```vb
Me.ctl.ForeShade=25
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)


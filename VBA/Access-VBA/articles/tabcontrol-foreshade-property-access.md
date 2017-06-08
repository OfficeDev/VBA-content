---
title: TabControl.ForeShade Property (Access)
keywords: vbaac10.chm14606
f1_keywords:
- vbaac10.chm14606
ms.prod: access
api_name:
- Access.TabControl.ForeShade
ms.assetid: a17b6a7e-d943-f785-459b-43a2c7998ff7
ms.date: 06/08/2017
---


# TabControl.ForeShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **ForeShade**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **ForeShade** property contains a numeric expression that can be used to darken the theme color in the **BackColor** property. The default value of the **ForeShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens  **ForeColor** by 75%.


```vb
Me.ctl.ForeShade=25
```


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)


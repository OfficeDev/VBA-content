---
title: Section.BackShade Property (Access)
keywords: vbaac10.chm14633
f1_keywords:
- vbaac10.chm14633
ms.prod: access
api_name:
- Access.Section.BackShade
ms.assetid: 36d9b31e-f632-cd7e-0ecf-6ea6de57e1a4
ms.date: 06/08/2017
---


# Section.BackShade Property (Access)

Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BackShade**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **BackShade** property contains a numeric expression that can be used to darken the theme color in the **BackColor** property. The default value of the **BackShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens the  **BackColor** property of the section by 75%.


```vb
Me.FormHeader.BackShade=25
```


## See also


#### Concepts


[Section Object](section-object-access.md)


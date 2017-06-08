---
title: CommandButton.BorderShade Property (Access)
keywords: vbaac10.chm14603
f1_keywords:
- vbaac10.chm14603
ms.prod: access
api_name:
- Access.CommandButton.BorderShade
ms.assetid: 52816058-36f4-3b68-38dd-5a1324b87428
ms.date: 06/08/2017
---


# CommandButton.BorderShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BorderShade**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **BorderShade** property contains a numeric expression that can be used to darken the theme color in the **BorderColor** property. The default value of the **BorderShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens the  **BorderColor** property by 75%.


```vb
Me.ctl.BorderShade=25
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)


---
title: NavigationButton.ForeTint Property (Access)
keywords: vbaac10.chm14605
f1_keywords:
- vbaac10.chm14605
ms.prod: access
api_name:
- Access.NavigationButton.ForeTint
ms.assetid: 4d19f8e8-67e0-412a-a515-dd357a26bc46
ms.date: 06/08/2017
---


# NavigationButton.ForeTint Property (Access)

Gets or sets the tint that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **ForeTint**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

The  **ForeTint** property contains a numeric expression that can be used to lighten the theme color in the **ForeColor** property. The default value of the **ForeTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example lightens the  **ForeColor** property by 75%.


```vb
Me.ctl.ForeTint=25
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)


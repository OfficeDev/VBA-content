---
title: BoundObjectFrame.BorderTint Property (Access)
keywords: vbaac10.chm14602
f1_keywords:
- vbaac10.chm14602
ms.prod: access
api_name:
- Access.BoundObjectFrame.BorderTint
ms.assetid: f171677a-d8a2-f0fb-233e-636ec13e20f8
ms.date: 06/08/2017
---


# BoundObjectFrame.BorderTint Property (Access)

Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BorderTint**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **BorderTint** property contains a numeric expression that can be used to lighten the theme color in the **BorderColor** property. The default value of the **BorderTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example lightens the  **BorderColor** property by 75%.


```vb
Me.ctl.BorderTint=25
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)


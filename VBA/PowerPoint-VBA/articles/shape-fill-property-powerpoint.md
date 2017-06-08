---
title: Shape.Fill Property (PowerPoint)
keywords: vbapp10.chm547022
f1_keywords:
- vbapp10.chm547022
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Fill
ms.assetid: bfb2dfe6-5036-0731-3a0f-1294ba87e103
ms.date: 06/08/2017
---


# Shape.Fill Property (PowerPoint)

Returns a  **[FillFormat](fillformat-object-powerpoint.md)** object that contains fill formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **Fill**

 _expression_ A variable that represents a **Shape** object.


### Return Value

FillFormat


## Example

This example adds a rectangle to  `myDocument` and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    .ForeColor.RGB = RGB(128, 0, 0)
    .BackColor.RGB = RGB(170, 170, 170)
    .TwoColorGradient msoGradientHorizontal, 1
End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)


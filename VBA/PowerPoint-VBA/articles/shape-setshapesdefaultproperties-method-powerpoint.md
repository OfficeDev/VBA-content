---
title: Shape.SetShapesDefaultProperties Method (PowerPoint)
keywords: vbapp10.chm547012
f1_keywords:
- vbapp10.chm547012
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.SetShapesDefaultProperties
ms.assetid: 4974cc1b-28af-94da-0821-76ffb698e2c3
ms.date: 06/08/2017
---


# Shape.SetShapesDefaultProperties Method (PowerPoint)

Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.


## Syntax

 _expression_. **SetShapesDefaultProperties**

 _expression_ A variable that represents a **Shape** object.


## Example

This example adds a rectangle to  `myDocument`, formats the rectangle's fill, applies the rectangle's formatting to the default shape, and then adds another smaller rectangle to the document. The second rectangle has the same fill as the first one.


```vb
Set mydocument = ActivePresentation.Slides(1)

With mydocument.Shapes

    With .AddShape(msoShapeRectangle, 5, 5, 80, 60)

        With .Fill

            .ForeColor.RGB = RGB(0, 0, 255)

            .BackColor.RGB = RGB(0, 204, 255)

            .Patterned msoPatternHorizontalBrick

        End With

    ' Sets formatting for default shapes

        .SetShapesDefaultProperties

    End With

' New shape has default formatting

    .AddShape msoShapeRectangle, 90, 90, 40, 30

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)


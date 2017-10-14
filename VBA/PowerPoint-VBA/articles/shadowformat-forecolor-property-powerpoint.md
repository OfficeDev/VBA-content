---
title: ShadowFormat.ForeColor Property (PowerPoint)
keywords: vbapp10.chm554004
f1_keywords:
- vbapp10.chm554004
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.ForeColor
ms.assetid: 8547b2b0-d9ba-41ce-9f13-659a8727cbdb
ms.date: 06/08/2017
---


# ShadowFormat.ForeColor Property (PowerPoint)

Returns or sets a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the foreground color for the fill, line, or shadow. Read/write.


## Syntax

 _expression_. **ForeColor**

 _expression_ A variable that represents a **ShadowFormat** object.


### Return Value

ColorFormat


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

This example adds a patterned line to  `myDocument`.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 100, 250, 0).Line

    .Weight = 6

    .ForeColor.RGB = RGB(0, 0, 255)

    .BackColor.RGB = RGB(128, 0, 0)

    .Pattern = msoPatternDarkDownwardDiagonal

End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)


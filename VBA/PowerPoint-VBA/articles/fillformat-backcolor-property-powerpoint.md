---
title: FillFormat.BackColor Property (PowerPoint)
keywords: vbapp10.chm552011
f1_keywords:
- vbapp10.chm552011
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.BackColor
ms.assetid: d78fa88b-578d-f469-f2e1-7564ebc91f8d
ms.date: 06/08/2017
---


# FillFormat.BackColor Property (PowerPoint)

Returns or sets a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the background color for the specified fill or patterned line. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

ColorFormat


## Example

This example adds a rectangle to  `myDocument` and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeRectangle, _
        90, 90, 90, 50).Fill
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


[FillFormat Object](fillformat-object-powerpoint.md)


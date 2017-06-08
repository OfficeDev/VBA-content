---
title: ShadowFormat Object (PowerPoint)
keywords: vbapp10.chm554000
f1_keywords:
- vbapp10.chm554000
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat
ms.assetid: 0bf08db8-2e44-4fc3-7f48-6017af881f72
ms.date: 06/08/2017
---


# ShadowFormat Object (PowerPoint)

Represents shadow formatting for a shape.


## Example

Use the  **Shadow** property to return a **ShadowFormat** object. The following example adds a shadowed rectangle to `myDocument`. The semitransparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeRectangle, _
        50, 50, 100, 200).Shadow

    .ForeColor.RGB = RGB(0, 0, 128)
    .OffsetX = 5
    .OffsetY = -3
    .Transparency = 0.5
    .Visible = True

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


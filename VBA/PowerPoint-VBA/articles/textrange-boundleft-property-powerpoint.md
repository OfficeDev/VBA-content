---
title: TextRange.BoundLeft Property (PowerPoint)
keywords: vbapp10.chm569006
f1_keywords:
- vbapp10.chm569006
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.BoundLeft
ms.assetid: 2641e084-6b6e-ff6e-c6a6-27cb84cbd4dd
ms.date: 06/08/2017
---


# TextRange.BoundLeft Property (PowerPoint)

Returns the distance (in points) from the left edge of the text bounding box for the specified text frame to the left edge of the slide. Read-only.


## Syntax

 _expression_. **BoundLeft**

 _expression_ A variable that represents a **TextRange** object.


### Return Value

Single


## Example

This example adds a rounded rectangle to slide one in the active presentation. The rectangle has the same dimensions as the text bounding box for shape one.


```vb
With Application.ActivePresentation.Slides(1).Shapes
    Set tr = .Item(1).TextFrame.TextRange
    Set roundRect = .AddShape(msoShapeRoundedRectangle, _
        tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
End With

With roundRect.Fill
    .ForeColor.RGB = RGB(255, 0, 128)
    .Transparency = 0.75
End With
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)


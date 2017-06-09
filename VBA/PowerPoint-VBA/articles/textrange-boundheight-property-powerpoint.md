---
title: TextRange.BoundHeight Property (PowerPoint)
keywords: vbapp10.chm569009
f1_keywords:
- vbapp10.chm569009
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.BoundHeight
ms.assetid: 8f3b9947-5ee3-260d-3d44-0ad2da422724
ms.date: 06/08/2017
---


# TextRange.BoundHeight Property (PowerPoint)

Returns the height (in points) of the text bounding box for the specified text frame. Read-only.


## Syntax

 _expression_. **BoundHeight**

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


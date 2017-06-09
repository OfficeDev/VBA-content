---
title: TextRange.BoundTop Property (PowerPoint)
keywords: vbapp10.chm569007
f1_keywords:
- vbapp10.chm569007
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.BoundTop
ms.assetid: cfc3baec-06c4-da2f-a233-afcb5301302a
ms.date: 06/08/2017
---


# TextRange.BoundTop Property (PowerPoint)

Returns the distance (in points) from the top of the of the text bounding box for the specified text frame to the top of the slide. Read-only.


## Syntax

 _expression_. **BoundTop**

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


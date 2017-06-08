---
title: TextFrame.MarginLeft Property (PowerPoint)
keywords: vbapp10.chm558003
f1_keywords:
- vbapp10.chm558003
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.MarginLeft
ms.assetid: c00a6b6c-0a67-5738-f31f-3714e2bf430d
ms.date: 06/08/2017
---


# TextFrame.MarginLeft Property (PowerPoint)

Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

 _expression_. **MarginLeft**

 _expression_ A variable that represents a **TextFrame** object.


### Return Value

Single


## Example

This example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeRectangle, _
        0, 0, 250, 140).TextFrame
    .TextRange.Text = "Here is some test text"
    .MarginBottom = 0
    .MarginLeft = 10
    .MarginRight = 0
    .MarginTop = 20
End With


```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)


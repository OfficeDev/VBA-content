---
title: TextFrame.MarginRight Property (PowerPoint)
keywords: vbapp10.chm558004
f1_keywords:
- vbapp10.chm558004
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.MarginRight
ms.assetid: 57ab53e7-1fbf-09b6-13c4-7cb0a814d9e3
ms.date: 06/08/2017
---


# TextFrame.MarginRight Property (PowerPoint)

Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

 _expression_. **MarginRight**

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
    .MarginRight = 5
    .MarginTop = 20
End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)


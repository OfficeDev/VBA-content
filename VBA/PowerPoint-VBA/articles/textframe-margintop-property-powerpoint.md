---
title: TextFrame.MarginTop Property (PowerPoint)
keywords: vbapp10.chm558005
f1_keywords:
- vbapp10.chm558005
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.MarginTop
ms.assetid: 78ae54cd-1841-950b-c06e-c693fa5daebb
ms.date: 06/08/2017
---


# TextFrame.MarginTop Property (PowerPoint)

Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

 _expression_. **MarginTop**

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


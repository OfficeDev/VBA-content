---
title: TextFrame.MarginBottom Property (PowerPoint)
keywords: vbapp10.chm558002
f1_keywords:
- vbapp10.chm558002
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.MarginBottom
ms.assetid: c1798b95-cb98-9dfd-5958-39fdea177b6e
ms.date: 06/08/2017
---


# TextFrame.MarginBottom Property (PowerPoint)

Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

 _expression_. **MarginBottom**

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


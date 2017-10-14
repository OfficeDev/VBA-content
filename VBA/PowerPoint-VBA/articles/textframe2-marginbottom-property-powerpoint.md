---
title: TextFrame2.MarginBottom Property (PowerPoint)
keywords: vbapp10.chm678002
f1_keywords:
- vbapp10.chm678002
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.MarginBottom
ms.assetid: f1a061e8-8248-9cbe-b4a7-09969644e5c0
ms.date: 06/08/2017
---


# TextFrame2.MarginBottom Property (PowerPoint)

Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

 _expression_. **MarginBottom**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

Single


## Example

The following example adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Public Sub MarginBottom_Example()



    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2

        .TextRange.Text = "Here is some sample text"

        .MarginBottom = 10

        .MarginLeft = 10

        .MarginRight = 10

        .MarginTop = 10

    End With

    

End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)


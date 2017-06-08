---
title: TextFrame2.MarginRight Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.MarginRight
ms.assetid: 82f3bd91-5250-b627-1a3a-780da3c9fc66
ms.date: 06/08/2017
---


# TextFrame2.MarginRight Property (Office)

Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write


## Syntax

 _expression_. **MarginRight**

 _expression_ An expression that returns a **TextFrame2** object.


## Example

The following example adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```
Set pptSlide = ActivePresentation.Slides(1) 
With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some sample text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)
#### Other resources


[TextFrame2 Object Members](textframe2-members-office.md)


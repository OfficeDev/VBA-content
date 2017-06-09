---
title: Shape.TextFrame Property (PowerPoint)
keywords: vbapp10.chm547035
f1_keywords:
- vbapp10.chm547035
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.TextFrame
ms.assetid: 6e4ad91e-c356-6a73-883d-8a0fd18c6ff6
ms.date: 06/08/2017
---


# Shape.TextFrame Property (PowerPoint)

Returns a  **[TextFrame](textframe-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape or master text style.


## Syntax

 _expression_. **TextFrame**

 _expression_ A variable that represents a **Shape** object.


### Return Value

TextFrame


## Remarks

Use the  **TextRange** property of the **TextFrame** object to return the text in the text frame.

Use the  **HasTextFrame** property to determine whether a shape contains a text frame before you apply the **TextFrame** property.


## Example

This example adds a rectangle to  `myDocument`, adds text to the rectangle, and sets the top margin for the text frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame
    .TextRange.Text = "Here is some test text"
    .MarginTop = 10
End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)


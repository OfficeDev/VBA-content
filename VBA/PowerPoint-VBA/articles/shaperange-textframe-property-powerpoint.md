---
title: ShapeRange.TextFrame Property (PowerPoint)
keywords: vbapp10.chm548035
f1_keywords:
- vbapp10.chm548035
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.TextFrame
ms.assetid: ec6093f2-232b-361b-b85d-7a99fafd8878
ms.date: 06/08/2017
---


# ShapeRange.TextFrame Property (PowerPoint)

Returns a  **[TextFrame](textframe-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape or master text style. Read-only.


## Syntax

 _expression_. **TextFrame**

 _expression_ A variable that represents a **ShapeRange** object.


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


[ShapeRange Object](shaperange-object-powerpoint.md)


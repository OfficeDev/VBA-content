---
title: TextRange2 Object (PowerPoint)
ms.assetid: 88e2de08-3d15-406d-99a0-93c3cd661eda
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2 Object (PowerPoint)

Represents the text frame in a  **Shape** or **ShapeRange** objects.


## Remarks

This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the  **TextFrame2** property to return a **TextFrame2** object.


## Example

The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame. 


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With 

```


## See also


#### Concepts







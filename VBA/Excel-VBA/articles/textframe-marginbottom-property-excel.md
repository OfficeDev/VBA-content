---
title: TextFrame.MarginBottom Property (Excel)
keywords: vbaxl10.chm644073
f1_keywords:
- vbaxl10.chm644073
ms.prod: excel
api_name:
- Excel.TextFrame.MarginBottom
ms.assetid: 50b2d7a9-8595-44a7-6f7a-ee2d2255869c
ms.date: 06/08/2017
---


# TextFrame.MarginBottom Property (Excel)

Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .


## Syntax

 _expression_ . **MarginBottom**

 _expression_ A variable that represents a **TextFrame** object.


## Example

This example adds a rectangle to  `myDocument`, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .Characters.Text = "Here is some test text" 
 .MarginBottom = 0 
 .MarginLeft = 100 
 .MarginRight = 0 
 .MarginTop = 20 
End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-excel.md)


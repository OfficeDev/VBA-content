---
title: TextFrame.MarginRight Property (Word)
keywords: vbawd10.chm162660454
f1_keywords:
- vbawd10.chm162660454
ms.prod: word
api_name:
- Word.TextFrame.MarginRight
ms.assetid: 9c59758e-8813-a035-b001-5eb57371e7fd
ms.date: 06/08/2017
---


# TextFrame.MarginRight Property (Word)

Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .


## Syntax

 _expression_ . **MarginRight**

 _expression_ An expression that returns a **[TextFrame](textframe-object-word.md)** object.


## Example

This example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 0 
 .MarginLeft = 100 
 .MarginRight = 0 
 .MarginTop = 20 
End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)


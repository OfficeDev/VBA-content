---
title: TextFrame.MarginBottom Property (Word)
keywords: vbawd10.chm162660452
f1_keywords:
- vbawd10.chm162660452
ms.prod: word
api_name:
- Word.TextFrame.MarginBottom
ms.assetid: 16e2f8ef-d28b-c61c-8a82-25c18c1252e0
ms.date: 06/08/2017
---


# TextFrame.MarginBottom Property (Word)

Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .


## Syntax

 _expression_ . **MarginBottom**

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


---
title: Shape.TextFrame Property (Word)
keywords: vbawd10.chm161480825
f1_keywords:
- vbawd10.chm161480825
ms.prod: word
api_name:
- Word.Shape.TextFrame
ms.assetid: c9ee1782-ecee-e83b-2014-62d0509237b7
ms.date: 06/08/2017
---


# Shape.TextFrame Property (Word)

Returns a  **TextFrame** object that contains the text for the specified shape.


## Syntax

 _expression_ . **TextFrame**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example adds a rectangle to  _myDocument_ , adds text to the rectangle, and sets the margins for the text frame.


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


[Shape Object](shape-object-word.md)


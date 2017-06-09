---
title: Shape.Fill Property (Word)
keywords: vbawd10.chm161480811
f1_keywords:
- vbawd10.chm161480811
ms.prod: word
api_name:
- Word.Shape.Fill
ms.assetid: 99a4d4f1-cc25-3b84-29ed-6e77a9a36765
ms.date: 06/08/2017
---


# Shape.Fill Property (Word)

Returns a  **[FillFormat](fillformat-object-word.md)** object that contains fill formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **Fill**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example adds a rectangle to myDocument and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = Documents(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)


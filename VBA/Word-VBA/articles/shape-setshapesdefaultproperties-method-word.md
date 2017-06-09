---
title: Shape.SetShapesDefaultProperties Method (Word)
keywords: vbawd10.chm161480726
f1_keywords:
- vbawd10.chm161480726
ms.prod: word
api_name:
- Word.Shape.SetShapesDefaultProperties
ms.assetid: 372bf936-720a-bb15-a7cc-0bb8ca20181d
ms.date: 06/08/2017
---


# Shape.SetShapesDefaultProperties Method (Word)

Applies the formatting of the default shape for a document to the specified shape.


## Syntax

 _expression_ . **SetShapesDefaultProperties**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

New shapes inherit many of their attributes from the default shape.


## Example

This example adds a rectangle to  _myDocument_ , formats the rectangle's fill, applies the rectangle's formatting to the default shape, and then adds another (smaller) rectangle to the document. The second rectangle has the same fill as the first one.


```vb
Set mydocument = ActiveDocument 
With mydocument.Shapes 
 With .AddShape(msoShapeRectangle, 5, 5, 80, 60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned msoPatternHorizontalBrick 
 End With 
 ' Sets formatting for default shapes 
 .SetShapesDefaultProperties 
 End With 
 ' New shape has default formatting 
 .AddShape msoShapeRectangle, 90, 90, 40, 30 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)


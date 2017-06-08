---
title: Selection.ChildShapeRange Property (Word)
keywords: vbawd10.chm158663677
f1_keywords:
- vbawd10.chm158663677
ms.prod: word
api_name:
- Word.Selection.ChildShapeRange
ms.assetid: 1b7c1010-19e1-e849-0040-70e231aac133
ms.date: 06/08/2017
---


# Selection.ChildShapeRange Property (Word)

Returns a  **[ShapeRange](shaperange-object-word.md)** collection representing the child shapes contained within a selection.


## Syntax

 _expression_ . **ChildShapeRange**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Example

This example creates a new document with a drawing canvas, populates the drawing canvas with shapes, and then, after checking that the shapes selected are child shapes, fills the child shapes with a pattern.


```vb
Sub ChildShapes() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 'Create a new document with a drawing canvas and shapes 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=100, Top:=100, Width:=200, Height:=200) 
 shpCanvas.CanvasItems.AddShape msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeOval, _ 
 Left:=0, Top:=50, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeDiamond, _ 
 Left:=0, Top:=100, Width:=100, Height:=100 
 
 'Select all shapes in the canvas 
 shpCanvas.CanvasItems.SelectAll 
 
 'Fill canvas child shapes with a pattern 
 If Selection.HasChildShapeRange = True Then 
 Selection.ChildShapeRange.Fill.Patterned msoPatternDivot 
 Else 
 MsgBox "This is not a range of child shapes." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)


---
title: Selection.ChildShapeRange Property (PowerPoint)
keywords: vbapp10.chm508011
f1_keywords:
- vbapp10.chm508011
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.ChildShapeRange
ms.assetid: f7458e07-47ec-c832-0731-94f4ba94ca89
ms.date: 06/08/2017
---


# Selection.ChildShapeRange Property (PowerPoint)

Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the child shapes of a selection.


## Syntax

 _expression_. **ChildShapeRange**

 _expression_ A variable that represents a **Selection** object.


### Return Value

ShapeRange


## Example

This example creates a new document with a drawing canvas, populates the drawing canvas with shapes, and selects the shapes added to the canvas. Then after checking that the shapes selected are child shapes, it fills the child shapes with a pattern.


```vb
Sub ChildShapes()

    Dim sldNew As Slide
    Dim shpCanvas As Shape

    'Create a new slide with a drawing canvas and shapes
    Set sldNew = Presentations(1).Slides _
        .Add(Index:=1, Layout:=ppLayoutBlank)

    Set shpCanvas = sldNew.Shapes.AddCanvas( _
        Left:=100, Top:=100, Width:=200, Height:=200)

    With shpCanvas.CanvasItems
        .AddShape msoShapeRectangle, Left:=0, Top:=0, _
            Width:=100, Height:=100
        .AddShape msoShapeOval, Left:=0, Top:=50, _
            Width:=100, Height:=100
        .AddShape msoShapeDiamond, Left:=0, Top:=100, _
            Width:=100, Height:=100
    End With

    'Select all shapes in the canvas
    shpCanvas.CanvasItems.SelectAll

    'Fill canvas child shapes with a pattern
    With ActiveWindow.Selection
        If .HasChildShapeRange = True Then
            .ChildShapeRange.Fill.Patterned Pattern:=msoPatternDivot
        Else
            MsgBox "This is not a range of child shapes."
        End If
    End With
	
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-powerpoint.md)


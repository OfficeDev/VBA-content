---
title: Cell.Parent Property (PowerPoint)
keywords: vbapp10.chm628002
f1_keywords:
- vbapp10.chm628002
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Parent
ms.assetid: 45650dd8-b51d-68ec-d117-5ddb8e8c675f
ms.date: 06/08/2017
---


# Cell.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the  **Shape** object that contains the text.


```vb
Set myShapes = ActivePresentation.Slides(1).Shapes

With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _
        Top:=50, Width:=300, Height:=150).TextFrame
    .TextRange.Text = "Test text"
    .Parent.Rotation = 45
End With
```


## See also


#### Concepts


[Cell Object](cell-object-powerpoint.md)


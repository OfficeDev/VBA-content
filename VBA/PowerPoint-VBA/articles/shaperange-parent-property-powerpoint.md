---
title: ShapeRange.Parent Property (PowerPoint)
keywords: vbapp10.chm548001
f1_keywords:
- vbapp10.chm548001
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Parent
ms.assetid: d43d43e8-8b92-bf87-fc4e-160166f26b10
ms.date: 06/08/2017
---


# ShapeRange.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **ShapeRange** object.


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


[ShapeRange Object](shaperange-object-powerpoint.md)


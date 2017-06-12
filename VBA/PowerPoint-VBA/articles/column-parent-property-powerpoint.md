---
title: Column.Parent Property (PowerPoint)
keywords: vbapp10.chm624002
f1_keywords:
- vbapp10.chm624002
ms.prod: powerpoint
api_name:
- PowerPoint.Column.Parent
ms.assetid: bd4c1a9b-5395-e881-912c-92ecbaa82a5c
ms.date: 06/08/2017
---


# Column.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **Column** object.


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


[Column Object](column-object-powerpoint.md)


---
title: PrintOptions.Parent Property (PowerPoint)
keywords: vbapp10.chm517008
f1_keywords:
- vbapp10.chm517008
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.Parent
ms.assetid: d93f28db-9198-2edb-d6d0-77484bc54c13
ms.date: 06/08/2017
---


# PrintOptions.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **PrintOptions** object.


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


[PrintOptions Object](printoptions-object-powerpoint.md)


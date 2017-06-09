---
title: ScaleEffect.Parent Property (PowerPoint)
keywords: vbapp10.chm660002
f1_keywords:
- vbapp10.chm660002
ms.prod: powerpoint
api_name:
- PowerPoint.ScaleEffect.Parent
ms.assetid: d95ae142-5fd5-114f-a200-6a7d23b0b2fd
ms.date: 06/08/2017
---


# ScaleEffect.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **ScaleEffect** object.


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


[ScaleEffect Object](scaleeffect-object-powerpoint.md)


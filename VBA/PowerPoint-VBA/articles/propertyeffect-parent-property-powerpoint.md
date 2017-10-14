---
title: PropertyEffect.Parent Property (PowerPoint)
keywords: vbapp10.chm662002
f1_keywords:
- vbapp10.chm662002
ms.prod: powerpoint
api_name:
- PowerPoint.PropertyEffect.Parent
ms.assetid: a6fa3de5-1c8a-cb3a-c47e-b41510834516
ms.date: 06/08/2017
---


# PropertyEffect.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **PropertyEffect** object.


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


[PropertyEffect Object](propertyeffect-object-powerpoint.md)


---
title: AnimationPoint.Parent Property (PowerPoint)
keywords: vbapp10.chm664002
f1_keywords:
- vbapp10.chm664002
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint.Parent
ms.assetid: e789fe23-b350-1a9c-0093-e6a9230f22a7
ms.date: 06/08/2017
---


# AnimationPoint.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents an **AnimationPoint** object.


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


[AnimationPoint Object](animationpoint-object-powerpoint.md)


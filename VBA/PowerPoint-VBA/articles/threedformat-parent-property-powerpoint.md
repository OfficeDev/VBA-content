---
title: ThreeDFormat.Parent Property (PowerPoint)
keywords: vbapp10.chm557001
f1_keywords:
- vbapp10.chm557001
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.Parent
ms.assetid: 558d1ae3-6d40-a13b-406e-d5e322938316
ms.date: 06/08/2017
---


# ThreeDFormat.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **ThreeDFormat** object.


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


[ThreeDFormat Object](threedformat-object-powerpoint.md)


---
title: ObjectVerbs.Parent Property (PowerPoint)
keywords: vbapp10.chm564002
f1_keywords:
- vbapp10.chm564002
ms.prod: powerpoint
api_name:
- PowerPoint.ObjectVerbs.Parent
ms.assetid: 6fdcb62a-4387-2af9-d146-3a74295685c4
ms.date: 06/08/2017
---


# ObjectVerbs.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents an **ObjectVerbs** object.


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


[ObjectVerbs Object](objectverbs-object-powerpoint.md)


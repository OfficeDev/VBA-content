---
title: NamedSlideShow.Parent Property (PowerPoint)
keywords: vbapp10.chm516002
f1_keywords:
- vbapp10.chm516002
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow.Parent
ms.assetid: e4c06441-b641-30a3-5eef-b6cbacfcb9e2
ms.date: 06/08/2017
---


# NamedSlideShow.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **NamedSlideShow** object.


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


[NamedSlideShow Object](namedslideshow-object-powerpoint.md)


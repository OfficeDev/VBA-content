---
title: ColorFormat.Parent Property (PowerPoint)
keywords: vbapp10.chm506001
f1_keywords:
- vbapp10.chm506001
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.Parent
ms.assetid: aa54269d-ee71-0a4e-7063-8631577d6e76
ms.date: 06/08/2017
---


# ColorFormat.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **ColorFormat** object.


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


[ColorFormat Object](colorformat-object-powerpoint.md)


---
title: PlaySettings.Parent Property (PowerPoint)
keywords: vbapp10.chm568002
f1_keywords:
- vbapp10.chm568002
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.Parent
ms.assetid: 88c43d67-7936-58b1-f5b2-22fea54de0bc
ms.date: 06/08/2017
---


# PlaySettings.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **PlaySettings** object.


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


[PlaySettings Object](playsettings-object-powerpoint.md)


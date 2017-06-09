---
title: TextStyle.Parent Property (PowerPoint)
keywords: vbapp10.chm579002
f1_keywords:
- vbapp10.chm579002
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyle.Parent
ms.assetid: 4b9be0da-adf7-eb57-e3b6-8df1d72684b3
ms.date: 06/08/2017
---


# TextStyle.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **TextStyle** object.


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


[TextStyle Object](textstyle-object-powerpoint.md)


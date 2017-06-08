---
title: Table.Parent Property (PowerPoint)
keywords: vbapp10.chm622002
f1_keywords:
- vbapp10.chm622002
ms.prod: powerpoint
api_name:
- PowerPoint.Table.Parent
ms.assetid: 1c873300-6a8d-bdd7-ff69-aa0ffc9aa157
ms.date: 06/08/2017
---


# Table.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **Table** object.


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


[Table Object](table-object-powerpoint.md)


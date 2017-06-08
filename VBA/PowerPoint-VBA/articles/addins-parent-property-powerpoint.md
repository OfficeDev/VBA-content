---
title: AddIns.Parent Property (PowerPoint)
keywords: vbapp10.chm520002
f1_keywords:
- vbapp10.chm520002
ms.prod: powerpoint
api_name:
- PowerPoint.AddIns.Parent
ms.assetid: 1669a388-0593-b968-efb4-e96c2c4b97a1
ms.date: 06/08/2017
---


# AddIns.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents an **AddIns** object.


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


[AddIns Object](addins-object-powerpoint.md)


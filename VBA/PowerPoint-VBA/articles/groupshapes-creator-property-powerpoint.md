---
title: GroupShapes.Creator Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.GroupShapes.Creator
ms.assetid: 89ba86e8-37fe-b30f-7dd6-a8dd8bf7b3b0
ms.date: 06/08/2017
---


# GroupShapes.Creator Property (PowerPoint)

Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.


## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **GroupShapes** object.


### Return Value

Long


## Remarks

The  **Creator** property is designed to be used in Microsoft Office applications for the Macintosh.


## Example

This example displays a message about the creator of myObject.


```vb
Set myObject = Application.ActivePresentation.Slides(1).Shapes(1)

If myObject.Creator = &;h50575054 Then

    MsgBox "This is a PowerPoint object"

Else

    MsgBox "This is not a PowerPoint object"

End If
```


## See also


#### Concepts


[GroupShapes Object](groupshapes-object-powerpoint.md)


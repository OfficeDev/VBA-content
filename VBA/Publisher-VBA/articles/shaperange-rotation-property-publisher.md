---
title: ShapeRange.Rotation Property (Publisher)
keywords: vbapb10.chm2293830
f1_keywords:
- vbapb10.chm2293830
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Rotation
ms.assetid: 0239aaae-18c7-56ef-f2b1-82f82660370a
ms.date: 06/08/2017
---


# ShapeRange.Rotation Property (Publisher)

Returns or sets a  **Single** that represents the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write.


## Syntax

 _expression_. **Rotation**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

To set the rotation of a three-dimensional shape around the x-axis or the y-axis, use the  **[RotationX](threedformat-rotationx-property-publisher.md)** property or the  **[RotationY](threedformat-rotationy-property-publisher.md)** property of the  **[ThreeDFormat](threedformat-object-publisher.md)** object.


## Example

This example matches the rotation of all shapes on the first page of the active publication to the rotation of the first shape. This example assumes there are at least two shapes on the first page of the active publication.


```vb
Sub SetShapeRotation() 
 Dim sngRotation As Single 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 sngRotation = .Item(1).Rotation 
 For intCount = 1 To .Count 
 .Item(intCount).Rotation = sngRotation 
 Next intCount 
 End With 
End Sub
```



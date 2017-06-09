---
title: Shape.Rotation Property (Word)
keywords: vbawd10.chm161480821
f1_keywords:
- vbawd10.chm161480821
ms.prod: word
api_name:
- Word.Shape.Rotation
ms.assetid: 7a66bdd7-ffda-64f2-8228-c1bce6d0640b
ms.date: 06/08/2017
---


# Shape.Rotation Property (Word)

Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write  **Single** .


## Syntax

 _expression_ . **Rotation**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

To set the rotation of a three-dimensional shape around the x-axis or the y-axis, use the  **RotationX** property or the **RotationY** property of the **ThreeDFormat** object.


## Example

This example matches the rotation of all shapes on myDocument to the rotation of shape one.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 sh1Rotation = .Item(1).Rotation 
 For o = 1 To .Count 
 .Item(o).Rotation = sh1Rotation 
 Next 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)


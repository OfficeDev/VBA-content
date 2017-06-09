---
title: ThreeDFormat.ResetRotation Method (PowerPoint)
keywords: vbapp10.chm557004
f1_keywords:
- vbapp10.chm557004
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.ResetRotation
ms.assetid: a766095a-f7a4-0fdf-8533-3ed00755950f
ms.date: 06/08/2017
---


# ThreeDFormat.ResetRotation Method (PowerPoint)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.


## Syntax

 _expression_. **ResetRotation**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the [RotationX](threedformat-rotationx-property-powerpoint.md)and  **[RotationY](threedformat-rotationy-property-powerpoint.md)** properties of the **ThreeDFormat** object. To set the extrusion rotation around the z-axis, use the **[Rotation](shape-rotation-property-powerpoint.md)** property of the **[Shape](shape-object-powerpoint.md)** object that represents the extruded shape.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of shape one on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).ThreeD.ResetRotation
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)


---
title: ThreeDFormat.ResetRotation Method (Excel)
keywords: vbaxl10.chm119002
f1_keywords:
- vbaxl10.chm119002
ms.prod: excel
api_name:
- Excel.ThreeDFormat.ResetRotation
ms.assetid: 55173d20-2d13-d3a8-39db-6b1a161c6ea6
ms.date: 06/08/2017
---


# ThreeDFormat.ResetRotation Method (Excel)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.


## Syntax

 _expression_ . **ResetRotation**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the  **[RotationX](threedformat-rotationx-property-excel.md)** and **[RotationY](threedformat-rotationy-property-excel.md)** properties of the **[ThreeDFormat](threedformat-object-excel.md)** object. To set the extrusion rotation around the z-axis, use the **[Rotation](shape-rotation-property-excel.md)** property of the **[Shape](shape-object-excel.md)** object that represents the extruded shape.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of shape one on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).ThreeD.ResetRotation
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)


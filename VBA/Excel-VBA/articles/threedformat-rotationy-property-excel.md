---
title: ThreeDFormat.RotationY Property (Excel)
keywords: vbaxl10.chm119015
f1_keywords:
- vbaxl10.chm119015
ms.prod: excel
api_name:
- Excel.ThreeDFormat.RotationY
ms.assetid: 71d6e255-eb1c-62bc-61f2-8b4f8be3ad6f
ms.date: 06/08/2017
---


# ThreeDFormat.RotationY Property (Excel)

Returns or sets the rotation of the extruded shape around the y-axis in degrees. Can be a value from - 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write  **Single** .


## Syntax

 _expression_ . **RotationY**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

To set the rotation of the extruded shape around the x-axis, use the  **[RotationX](threedformat-rotationx-property-excel.md)** property of the **ThreeDFormat** object. To set the rotation of the extruded shape around the z-axis, use the **[Rotation](shape-rotation-property-excel.md)** property of the **[Shape](shape-object-excel.md)** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **[SetExtrusionDirection](threedformat-setextrusiondirection-method-excel.md)** method.


## Example

This example adds three identical extruded ovals to  `myDocument` and sets their rotation around the y-axis to - 30, 0, and 30 degrees, respectively.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 With .AddShape(msoShapeOval, 30, 30, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = -30 
 End With 
 With .AddShape(msoShapeOval, 30, 70, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = 0 
 End With 
 With .AddShape(msoShapeOval, 30, 110, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = 30 
 End With 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)


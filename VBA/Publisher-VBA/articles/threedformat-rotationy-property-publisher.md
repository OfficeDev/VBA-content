---
title: ThreeDFormat.RotationY Property (Publisher)
keywords: vbapb10.chm3801360
f1_keywords:
- vbapb10.chm3801360
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.RotationY
ms.assetid: 571f090b-71a8-c92e-b4d8-4f21a4c383ed
ms.date: 06/08/2017
---


# ThreeDFormat.RotationY Property (Publisher)

Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Can be a value from - 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write  **Single**.


## Syntax

 _expression_. **RotationY**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

Single


## Remarks

To set the rotation of the extruded shape around the x-axis, use the  **[RotationX](threedformat-rotationx-property-publisher.md)** property of the  **ThreeDFormat** object. To set the rotation of the extruded shape around the z-axis, use the **[Rotation](shape-rotation-property-publisher.md)** property of the  **[Shape](shape-object-publisher.md)** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** method.


## Example

This example adds three identical extruded ovals to the active document and sets their rotation around the y-axis to - 30, 0, and 30 degrees, respectively.


```vb
Sub SetRotationY() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShapeOval, Left:=30, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = -30 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=90, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = 0 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=150, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = 30 
 End With 
 End With 
End Sub
```



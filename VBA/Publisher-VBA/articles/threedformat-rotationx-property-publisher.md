---
title: ThreeDFormat.RotationX Property (Publisher)
keywords: vbapb10.chm3801353
f1_keywords:
- vbapb10.chm3801353
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.RotationX
ms.assetid: 1ee394cb-746b-02f0-f2af-aa4a6fffd172
ms.date: 06/08/2017
---


# ThreeDFormat.RotationX Property (Publisher)

Returns or sets the rotation of the extruded shape around the x-axis in degrees. Can be a value from - 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write  **Single**.


## Syntax

 _expression_. **RotationX**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

Single


## Remarks

To set the rotation of the extruded shape around the y-axis, use the  **[RotationY](threedformat-rotationy-property-publisher.md)** property of the  **ThreeDFormat** object. To set the rotation of the extruded shape around the z-axis, use the **[Rotation](shape-rotation-property-publisher.md)** property of the  **[Shape](shape-object-publisher.md)** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** method.


## Example

This example adds three identical extruded ovals to the active document and sets their rotation around the x-axis to - 30, 0, and 30 degrees, respectively.


```vb
Sub SetRotationX() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShapeOval, Left:=30, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = -30 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=90, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 0 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=150, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 30 
 End With 
 End With 
End Sub
```



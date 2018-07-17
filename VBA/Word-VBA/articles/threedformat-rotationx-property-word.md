---
title: ThreeDFormat.RotationX Property (Word)
keywords: vbawd10.chm164626541
f1_keywords:
- vbawd10.chm164626541
ms.prod: word
api_name:
- Word.ThreeDFormat.RotationX
ms.assetid: 8ed5e2de-8a1b-e75e-da7d-10b6d1d1a988
ms.date: 06/08/2017
---


# ThreeDFormat.RotationX Property (Word)

Returns or sets the rotation of the extruded shape around the x-axis in degrees. Read/write  **Single** .


## Syntax

 _expression_ . **RotationX**

 _expression_ An expression that returns a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

The  **RotationX** property can be a value from - 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation.

To set the rotation of the extruded shape around the y-axis, use the  **[RotationY](threedformat-rotationy-property-word.md)** property of the ThreeDFormat object. To set the rotation of the extruded shape around the z-axis, use the **[Rotation](shape-rotation-property-word.md)** property of the **[Shape](shape-object-word.md)** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **[SetExtrusionDirection](threedformat-setextrusiondirection-method-word.md)** method.


## Example

This example adds three identical extruded ovals to the active document and sets their rotation around the x-axis to - 30, 0, and 30 degrees, respectively.


```vb
With ActiveDocument.Shapes 
 With .AddShape(msoShapeOval, 30, 60, 50, 25).ThreeD 
 .Visible = True 
 .RotationX = -30 
 End With 
 With .AddShape(msoShapeOval, 90, 60, 50, 25).ThreeD 
 .Visible = True 
 .RotationX = 0 
 End With 
 With .AddShape(msoShapeOval, 150, 60, 50, 25).ThreeD 
 .Visible = True 
 .RotationX = 30 
 End With 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)


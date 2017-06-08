---
title: ThreeDFormat.RotationX Property (PowerPoint)
keywords: vbapp10.chm557016
f1_keywords:
- vbapp10.chm557016
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.RotationX
ms.assetid: 8c434ef8-1364-5989-71da-e96ddfbd15ac
ms.date: 06/08/2017
---


# ThreeDFormat.RotationX Property (PowerPoint)

Returns or sets the rotation of the extruded shape around the x-axis, in degrees. Read/write.


## Syntax

 _expression_. **RotationX**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

Single


## Remarks

Can be a value from - 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation.

To set the rotation of the extruded shape around the y-axis, use the  **[RotationY](threedformat-rotationy-property-powerpoint.md)** property of the **ThreeDFormat** object.

To set the rotation of the extruded shape around the z-axis, use the  **[Rotation](shape-rotation-property-powerpoint.md)** property of the **[Shape](shape-object-powerpoint.md)** object.

To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-powerpoint.md)** method.


## Example

This example adds three identical extruded ovals to  `myDocument` and sets their rotation around the x-axis to - 30, 0, and 30 degrees, respectively.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

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


[ThreeDFormat Object](threedformat-object-powerpoint.md)


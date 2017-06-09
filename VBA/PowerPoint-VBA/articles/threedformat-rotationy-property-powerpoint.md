---
title: ThreeDFormat.RotationY Property (PowerPoint)
keywords: vbapp10.chm557017
f1_keywords:
- vbapp10.chm557017
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.RotationY
ms.assetid: 1e39544d-e95d-7419-8d6b-140696a43895
ms.date: 06/08/2017
---


# ThreeDFormat.RotationY Property (PowerPoint)

Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Read/write.


## Syntax

 _expression_. **RotationY**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

Single


## Remarks

Can be a value from - 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right.

To set the rotation of the extruded shape around the x-axis, use the  **[RotationX](threedformat-rotationx-property-powerpoint.md)** property of the **ThreeDFormat** object.

To set the rotation of the extruded shape around the z-axis, use the  **[Rotation](shape-rotation-property-powerpoint.md)** property of the **[Shape](shape-object-powerpoint.md)** object.

To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the  **[SetExtrusionDirection](threedformat-setextrusiondirection-method-powerpoint.md)** method.


## Example

This example adds three identical extruded ovals to  `myDocument` and sets their rotation around the y-axis to - 30, 0, and 30 degrees, respectively.


```vb
Set myDocument = ActivePresentation.Slides(1)

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


[ThreeDFormat Object](threedformat-object-powerpoint.md)


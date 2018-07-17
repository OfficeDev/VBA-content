---
title: ThreeDFormat.Depth Property (PowerPoint)
keywords: vbapp10.chm557007
f1_keywords:
- vbapp10.chm557007
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.Depth
ms.assetid: ef38cda1-5bf0-df3e-aee5-96f18fb1c600
ms.date: 06/08/2017
---


# ThreeDFormat.Depth Property (PowerPoint)

Returns or sets the depth of the shape's extrusion. Read/write.


## Syntax

 _expression_. **Depth**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

Single


## Remarks

The  **Depth** property value can be from - 600 through 9600 (positive values produce an extrusion whose front face is the original shape; negative values produce an extrusion whose back face is the original shape).


## Example

This example adds an oval to  `myDocument`, and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set myShape = myDocument.Shapes _
    .AddShape(msoShapeOval, 90, 90, 90, 40)

With myShape.ThreeD
    .Visible = True
    .Depth = 50
    'RGB value for purple
    .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-powerpoint.md)
[ThreeDFormat Object](threedformat-object-powerpoint.md)


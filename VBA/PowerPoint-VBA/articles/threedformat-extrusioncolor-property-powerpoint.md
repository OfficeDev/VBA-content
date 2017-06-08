---
title: ThreeDFormat.ExtrusionColor Property (PowerPoint)
keywords: vbapp10.chm557008
f1_keywords:
- vbapp10.chm557008
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.ExtrusionColor
ms.assetid: 70d290bd-84e6-8a24-0c87-3a9b10ae4282
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColor Property (PowerPoint)

Returns a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the color of the shape's extrusion. Read-only.


## Syntax

 _expression_. **ExtrusionColor**

 _expression_ A variable that represents an **ThreeDFormat** object.


### Return Value

ColorFormat


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


[ThreeDFormat Object](threedformat-object-powerpoint.md)


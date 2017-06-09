---
title: ThreeDFormat Object (PowerPoint)
keywords: vbapp10.chm557000
f1_keywords:
- vbapp10.chm557000
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat
ms.assetid: d6eb7b36-57df-727e-fc5b-50b8c4790c1c
ms.date: 06/08/2017
---


# ThreeDFormat Object (PowerPoint)

Represents a shape's three-dimensional formatting.


## Remarks

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the  **ThreeDFormat** object for such a shape will fail.


## Example

Use the  **ThreeD** property to return a **ThreeDFormat** object. The following example adds an oval to `myDocument` and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


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


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


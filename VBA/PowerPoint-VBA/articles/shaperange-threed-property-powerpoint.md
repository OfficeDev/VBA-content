---
title: ShapeRange.ThreeD Property (PowerPoint)
keywords: vbapp10.chm548036
f1_keywords:
- vbapp10.chm548036
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ThreeD
ms.assetid: e0e2f72d-639b-86fd-2191-f537ddcd45ad
ms.date: 06/08/2017
---


# ShapeRange.ThreeD Property (PowerPoint)

Returns a  **[ThreeDFormat](threedformat-object-powerpoint.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **ThreeD**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

ThreeDFormat


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3-D effects applied to shape one on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .Depth = 50

    'RGB value for purple

    .ExtrusionColor.RGB = RGB(255, 100, 255)

    .SetExtrusionDirection msoExtrusionTop

    .PresetLightingDirection = msoLightingLeft

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)


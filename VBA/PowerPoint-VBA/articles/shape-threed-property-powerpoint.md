---
title: Shape.ThreeD Property (PowerPoint)
keywords: vbapp10.chm547036
f1_keywords:
- vbapp10.chm547036
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ThreeD
ms.assetid: 16f0bc6a-ae6c-f4c3-9e3c-641f069eb7f6
ms.date: 06/08/2017
---


# Shape.ThreeD Property (PowerPoint)

Returns a  **[ThreeDFormat](threedformat-object-powerpoint.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **ThreeD**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)


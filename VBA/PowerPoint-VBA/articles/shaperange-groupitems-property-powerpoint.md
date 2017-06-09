---
title: ShapeRange.GroupItems Property (PowerPoint)
keywords: vbapp10.chm548023
f1_keywords:
- vbapp10.chm548023
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.GroupItems
ms.assetid: 94d0e684-5237-2415-e222-cd38cbd22e36
ms.date: 06/08/2017
---


# ShapeRange.GroupItems Property (PowerPoint)

Returns a  **[GroupShapes](groupshapes-object-powerpoint.md)** object that represents the individual shapes in the specified group. Use the **Item** method of the **GroupShapes** object to return a single shape from the group. Read-only.


## Syntax

 _expression_. **GroupItems**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

GroupShapes


## Example

This example adds three triangles to  `myDocument`, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    .AddShape(msoShapeIsoscelesTriangle, 10, _
        10, 100, 100).Name = "shpOne"

    .AddShape(msoShapeIsoscelesTriangle, 150, _
        10, 100, 100).Name = "shpTwo"

    .AddShape(msoShapeIsoscelesTriangle, 300, _
        10, 100, 100).Name = "shpThree"

    With .Range(Array("shpOne", "shpTwo", "shpThree")).Group
        .Fill.PresetTextured msoTextureBlueTissuePaper
        .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble
    End With
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)


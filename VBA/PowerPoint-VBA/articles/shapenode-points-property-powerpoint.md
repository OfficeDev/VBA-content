---
title: ShapeNode.Points Property (PowerPoint)
keywords: vbapp10.chm561003
f1_keywords:
- vbapp10.chm561003
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNode.Points
ms.assetid: 1ba61c2f-708d-d2a5-aac0-68f566f19337
ms.date: 06/08/2017
---


# ShapeNode.Points Property (PowerPoint)

Returns a  **Variant** that represents the position of the specified node as a coordinate pair. Read-only.


## Syntax

 _expression_. **Points**

 _expression_ A variable that represents a **ShapeNode** object.


### Return Value

Variant


## Remarks

Each coordinate is expressed in points. Use the  **[SetPosition](shapenodes-setposition-method-powerpoint.md)** method to set the value of this property.


## Example

This example moves node two in shape three in the active presentation to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```vb
With ActivePresentation.Slides(1).Shapes(3).Nodes

    pointsArray = .Item(2).Points

    currXvalue = pointsArray(1, 1)

    currYvalue = pointsArray(1, 2)

    .SetPosition Index:=2, X1:=currXvalue + 200, Y1:=currYvalue + 300

End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-powerpoint.md)


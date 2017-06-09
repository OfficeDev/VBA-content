---
title: ShapeRange.Group Method (PowerPoint)
keywords: vbapp10.chm548061
f1_keywords:
- vbapp10.chm548061
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Group
ms.assetid: f70f3986-3a39-78f9-476e-b72ef000c469
ms.date: 06/08/2017
---


# ShapeRange.Group Method (PowerPoint)

Groups the shapes in the specified range. Returns the grouped shapes as a single  **[Shape](shape-object-powerpoint.md)** object.


## Syntax

 _expression_. **Group**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Shape


## Remarks

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **[Shapes](shapes-object-powerpoint.md)** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example adds two shapes to  `myDocument`, groups the two new shapes, sets the fill for the group, rotates the group, and sends the group to the back of the drawing layer.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne"

    .AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"

    With .Range(Array("shpOne", "shpTwo")).Group

        .Fill.PresetTextured msoTextureBlueTissuePaper

        .Rotation = 45

        .ZOrder msoSendToBack

    End With

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)


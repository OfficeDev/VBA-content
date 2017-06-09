---
title: FreeformBuilder Object (PowerPoint)
keywords: vbapp10.chm546000
f1_keywords:
- vbapp10.chm546000
ms.prod: powerpoint
api_name:
- PowerPoint.FreeformBuilder
ms.assetid: fa188c8b-0781-dc9d-dd8d-3fc24c02d086
ms.date: 06/08/2017
---


# FreeformBuilder Object (PowerPoint)

Represents the geometry of a freeform while it is being built.


## Example

Use the [BuildFreeform](shapes-buildfreeform-method-powerpoint.md)method to return a  **FreeformBuilder** object. Use the[AddNodes](freeformbuilder-addnodes-method-powerpoint.md)method to add nodes to the freefrom. Use the [ConvertToShape](freeformbuilder-converttoshape-method-powerpoint.md)method to create the shape defined in the  **FreeformBuilder** object and add it to the **[Shapes](shapes-object-powerpoint.md)** collection. The following example adds a freeform with four segments to `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    .AddNodes msoSegmentCurve, msoEditingCorner, _
        380, 230, 400, 250, 450, 300
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 40
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
    .ConvertToShape
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


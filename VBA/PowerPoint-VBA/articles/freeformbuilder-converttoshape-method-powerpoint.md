---
title: FreeformBuilder.ConvertToShape Method (PowerPoint)
keywords: vbapp10.chm546003
f1_keywords:
- vbapp10.chm546003
ms.prod: powerpoint
api_name:
- PowerPoint.FreeformBuilder.ConvertToShape
ms.assetid: bc3d209e-6735-3011-9334-46049d269355
ms.date: 06/08/2017
---


# FreeformBuilder.ConvertToShape Method (PowerPoint)

Creates a shape that has the geometric characteristics of the specified  **[FreeformBuilder](freeformbuilder-object-powerpoint.md)** object. Returns a **[Shape](shape-object-powerpoint.md)** object that represents the new shape.


## Syntax

 _expression_. **ConvertToShape**

 _expression_ A variable that represents a **FreeformBuilder** object.


### Return Value

Shape


## Remarks

You must apply the [AddNodes](freeformbuilder-addnodes-method-powerpoint.md)method to a  **FreeformBuilder** object at least once before you use the **ConvertToShape** method.


## Example

This example adds a freeform with five vertices to the first slide in the active presentation.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    .AddNodes msoSegmentCurve, _
        msoEditingCorner, 380, 230, 400, 250, 450, 300
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 400
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
    .ConvertToShape
End With
```


## See also


#### Concepts


[FreeformBuilder Object](freeformbuilder-object-powerpoint.md)


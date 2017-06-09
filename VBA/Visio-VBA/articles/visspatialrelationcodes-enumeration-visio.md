---
title: VisSpatialRelationCodes Enumeration (Visio)
keywords: vis_sdr.chm70225
f1_keywords:
- vis_sdr.chm70225
ms.prod: visio
ms.assetid: 4834dcb7-48e4-14c4-272f-3531892a0ccd
ms.date: 06/08/2017
---


# VisSpatialRelationCodes Enumeration (Visio)

Codes for spatial relationships between shapes to be passed to the  **Shape.SpatialRelation** property.


## Remarks

The spatial relationship between shapes can be indicated by any combination of the following values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSpatialContainedIn**|&;H4|A shape can be contained within another shape. Shape B is contained within shape A if shape A encloses every region and path of shape B.|
| **visSpatialContain**|&;H2|A shape can contain another shape. Shape A contains shape B if shape A encloses every region and path of shape B.|
| **visSpatialOverlap**|&;H1|Two shapes can overlap. Shapes overlap if their interior regions have at least one point in common. You will also get this result if you compare a shape to itself or if either shape is a sub-shape of the other.|
| **visSpatialTouching**|&;H8|A shape can be touching another shape. Shape A touches shape B if neither one contains or overlaps the other and they have one or more common points whose distance is within the specified tolerance.|


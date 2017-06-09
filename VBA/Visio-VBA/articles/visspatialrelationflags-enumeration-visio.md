---
title: VisSpatialRelationFlags Enumeration (Visio)
keywords: vis_sdr.chm70230
f1_keywords:
- vis_sdr.chm70230
ms.prod: visio
ms.assetid: 38e44579-2e2c-cdb9-524b-e2b864901c13
ms.date: 06/08/2017
---


# VisSpatialRelationFlags Enumeration (Visio)

Flags passed to various properties of the  **Shape** object, including the **DistanceFrom** , **DistanceFromPoint** , **SpatialNeighbors** , **SpatialRelation** , and **SpatialSearch** properties.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSpatialBackToFront**|&;H8|Order items back to front.|
| **visSpatialFrontToBack**|&;H4|Order items front to back.|
| **visSpatialIgnoreVisible**|&;H20|Do not consider visible Geometry sections. By default, visible Geometry sections influence the result.|
| **visSpatialIncludeContainerShapes**|&;H80|Include containers. By default, containers are not included.|
| **visSpatialIncludeDataGraphics**|&;H40|Include data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.|
| **visSpatialIncludeGuides**|&;H2|Consider a guide's Geometry section. By default, guides do not influence the result.|
| **visSpatialIncludeHidden**|&;H10|Consider hidden Geometry sections. By default, hidden Geometry sections do not influence the result.|


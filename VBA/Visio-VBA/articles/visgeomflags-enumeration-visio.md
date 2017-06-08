---
title: VisGeomFlags Enumeration (Visio)
keywords: vis_sdr.chm70245
f1_keywords:
- vis_sdr.chm70245
ms.prod: visio
ms.assetid: 47462624-5d34-2643-66f7-bfde9eecbcce
ms.date: 06/08/2017
---


# VisGeomFlags Enumeration (Visio)

Flags to pass to methods of the  **Row** object that get and put vertex arrays, such as **GetPolylineData** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGeomExcludeLastPoint**|1 (&;H1)|The last point (the X and Y cells in the row) is not included in the data.|
| **visGeomWHPct**|16 (&;H10)|The X and Y values are percentages of width and height.|
| **visGeomXYLocal**|32 (&;H20)|The X and Y values are local, internal units in the drawing.|


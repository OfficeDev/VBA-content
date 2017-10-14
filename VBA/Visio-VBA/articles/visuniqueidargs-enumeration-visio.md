---
title: VisUniqueIDArgs Enumeration (Visio)
keywords: vis_sdr.chm70075
f1_keywords:
- vis_sdr.chm70075
ms.prod: visio
ms.assetid: 7268c074-3de9-72c8-d20e-1f6008aff347
ms.date: 06/08/2017
---


# VisUniqueIDArgs Enumeration (Visio)

Action codes to be passed to the  **Shape.UniqueID** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDeleteGUID**|2|Clear the unique ID of a shape and return a zero-length string ("").|
| **visDeleteGUIDWithUndo**|4|Clear the unique ID of a shape and return a zero-length string (""). Undoable.|
| **visGetGUID**|0|Return the unique ID string only if the shape already has a unique ID.|
| **visGetOrMakeGUID**|1|Return the unique ID string of the shape. If the shape does not already have a unique ID, assign one to the shape and return the new ID. |
| **visGetOrMakeGUIDWithUndo**|3|Return the unique ID string of the shape. If the shape does not already have a unique ID, assign one to the shape and return the new ID. Undoable.|


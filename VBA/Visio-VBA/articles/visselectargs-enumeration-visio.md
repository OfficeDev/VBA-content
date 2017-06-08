---
title: VisSelectArgs Enumeration (Visio)
keywords: vis_sdr.chm70070
f1_keywords:
- vis_sdr.chm70070
ms.prod: visio
ms.assetid: 21651fc7-c311-aefb-9f6c-27fcbf9740be
ms.date: 06/08/2017
---


# VisSelectArgs Enumeration (Visio)

Selection-type constants to be passed to the  **Selection.Select** and **Window.Select** methods.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDeselect**|1|Deselects a shape but leaves the rest of the selection unchanged.|
| **visSelect**|2|Selects a shape but leaves the rest of the selection unchanged.|
| **visSubSelect**|3|Selects a shape whose parent is already selected.|
| **visSelectAll**|4|Selects a shape and all its peers.|
| **visDeselectAll**|256|Deselects a shape and all its peers.|


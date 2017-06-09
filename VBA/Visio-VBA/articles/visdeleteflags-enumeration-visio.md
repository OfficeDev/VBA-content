---
title: VisDeleteFlags Enumeration (Visio)
keywords: vis_sdr.chm70670
f1_keywords:
- vis_sdr.chm70670
ms.prod: visio
api_name:
- Visio.VisDeleteFlags
ms.assetid: 1f36b2c8-1c46-519f-b0f0-b548363891ab
ms.date: 06/08/2017
---


# VisDeleteFlags Enumeration (Visio)

Specifies constants that define particular sets of instruction to apply to a deletion; passed to the  **[Selection.DeleteEx](selection-deleteex-method-visio.md)** and **[Shape.DeleteEx](shape-deleteex-method-visio.md)** methods.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDeleteNormal**|0|Match the deletion behavior in the user interface.|
| **visDeleteHealConnectors**|1|Delete connectors attached to deleted shapes.|
| **visDeleteNoHealConnectors**|2|Do not delete connectors attached to deleted shapes.|
| **visDeleteNoContainerMembers**|4|Do not delete unselected members of containers or lists.|
| **visDeleteNoAssociatedCallouts**|8|Do not delete unselected callouts associated with shapes.|


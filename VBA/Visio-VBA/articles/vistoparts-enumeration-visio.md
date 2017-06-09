---
title: VisToParts Enumeration (Visio)
keywords: vis_sdr.chm70165
f1_keywords:
- vis_sdr.chm70165
ms.prod: visio
ms.assetid: abf9c04f-b9aa-d6da-98f5-f3a293b2b0fd
ms.date: 06/08/2017
---


# VisToParts Enumeration (Visio)

Values returned by the  **Connect.ToPart** property.


## Remarks

The  **VisToParts** return codes indicate the part of a shape to which a connection is made.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visConnectionPoint**|100 + row index of connection point|Connect to specified connection point on target shape.|
| **visConnectToError**|-1|Error connecting to shape.|
| **visGuideIntersect**|4|Connect to intersection of guides on target shape.|
| **visGuideX**|1|Connect to vertical guide on target shape.|
| **visGuideY**|2|Connect to horizontal guide on target shape.|
| **visToAngle**|7|Connect to angle on target shape.|
| **visToNone**|0|Do not connect.|
| **visWholeShape**|3|Connect to entire target shape, using dynamic glue.|


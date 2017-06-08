---
title: VisConnectedShapesFlags Enumeration (Visio)
keywords: vis_sdr.chm70575
f1_keywords:
- vis_sdr.chm70575
ms.prod: visio
api_name:
- Visio.VisConnectedShapesFlags
ms.assetid: 00cf06f7-8161-8b56-9f3f-bed817789895
ms.date: 06/08/2017
---


# VisConnectedShapesFlags Enumeration (Visio)

Specifies constants that identify shapes by the directionality of their connectors; passed to the  **[Shape.ConnectedShapes](shape-connectedshapes-method-visio.md)** method.


 **Note**  Connection points that have dual directionality (both inward and outward) are identified as either incoming or outgoing based on how they are used in a particular connection.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visConnectedShapesAllNodes**|0|The shapes that are connected by using either incoming or outgoing connections.|
| **visConnectedShapesIncomingNodes**|1|The shapes that are connected by using incoming connections.|
| **visConnectedShapesOutgoingNodes**|2|The shapes that are connected by using outgoing connections.|


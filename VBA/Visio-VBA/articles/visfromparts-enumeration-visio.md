---
title: VisFromParts Enumeration (Visio)
keywords: vis_sdr.chm70160
f1_keywords:
- vis_sdr.chm70160
ms.prod: visio
ms.assetid: 243245c8-8683-1d7b-17cc-95691310537a
ms.date: 06/08/2017
---


# VisFromParts Enumeration (Visio)

Codes returned by the  **Connect.FromPart** property.


 **Note**  The  **visControlPoint** value is actually 100 plus the zero-based row index. For example, if the control point is in row 0, **visControlPoint** equals 100; if the control point is in row 1, **visControlPoint** equals 101.


## Remarks

The  **VisFromParts** return codes indicate the part of a shape from which a connection originates.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visBeginX**|7|Connection is from the begin point x of a 1-D shape.|
| **visBeginY**|8|Connection is from the begin point y of a 1-D shape.|
| **visBegin**|9|Connection is from the begin point of a 1-D shape.|
| **visBottomEdge**|4|Connection is from bottom edge of shape.|
| **visCenterEdge**|2|Connection is from the center (x) of a 1-D shape.|
| **visConnectFromError**|-1|Connection from an unknown part.|
| **visControlPoint**|100|Connection is from the control point plus the row index (see Note).|
| **visEndX**|10|Connection is from the endpoint (x) of a 1-D shape.|
| **visEndY**|11|Connection is from the endpoint (y) of a 1-D shape.|
| **visEnd**|12|Connection is from the end of a 1-D shape.|
| **visFromAngle**|13|Connection is from the direction of a connection point.|
| **visFromNone**|0|Connection is from nothing.|
| **visFromPin**|14|Connection is from the pin of a shape.|
| **visLeftEdge**|1|Connection is from the left edge of a shape.|
| **visMiddleEdge**|5|Connection is from the middle (y) of a shape.|
| **visRightEdge**|3|Connection is from the right edge of a shape.|
| **visTopEdge**|6|Connection is from the top edge of a shape.|


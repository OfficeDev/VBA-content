---
title: Curve Object (Visio)
keywords: vis_sdr.chm10075
f1_keywords:
- vis_sdr.chm10075
ms.prod: visio
api_name:
- Visio.Curve
ms.assetid: 040f47b2-794d-72c7-7479-b61d8f1cb75f
ms.date: 06/08/2017
---


# Curve Object (Visio)

An item in a  **Path** object that represents a consecutive sequence of rows in the Geometry section of its **Path** object.


## Remarks

The default property of  **Curve** object is **Point** .

If a  **Curve** object is in a collection returned by the **Paths** property of a **Shape** object, its coordinates are expressed in the shape's parent coordinate system. If the **Curve** object is in a collection returned by the **PathsLocal** property of a **Shape** object, its coordinates are expressed in the shape's local coordinate system. In both cases, the coordinates are expressed in internal drawing units (inches).

A  **Curve** object describes itself in terms of its parameter domain, which is the range [Start(),End()]. Use the **Start** property of a **Curve** object to obtain the curve's starting point and the **End** property of a **Curve** object to obtain the curve's ending point.

Use the  **Point** method of a curve object to extrapolate a point along the curve's path. Use the **PointAndDerivatives** method of a **Curve** object to determine a point along the curve's path and, optionally, its first and second derivatives.

Use the  **Points** property of a **Curve** object to obtain a stream of points that approximate the curve's path.



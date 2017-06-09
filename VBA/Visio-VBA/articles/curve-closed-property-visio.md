---
title: Curve.Closed Property (Visio)
keywords: vis_sdr.chm15513250
f1_keywords:
- vis_sdr.chm15513250
ms.prod: visio
api_name:
- Visio.Curve.Closed
ms.assetid: ed4a1f5c-c4e3-9da7-cfe0-4d42cc0dc6b5
ms.date: 06/08/2017
---


# Curve.Closed Property (Visio)

Determines if an object is closed (that is, if its begin point coincides with its endpoint). Read-only.


## Syntax

 _expression_ . **Closed**

 _expression_ A variable that represents a **Curve** object.


### Return Value

Integer


## Remarks

Use the  **Closed** property of a **Path** or **Curve** object to test for equality (Microsoft Visio uses 10E-6 as its "fuzz" factor) of the object's begin and endpoints. A closed **Curve** object can be in a **Path** object that is open, and a **Curve** object that is open can be in a closed **Path** object.

The  **Closed** property of a **Path** object is unrelated to a **Path** object's fill. A **Path** object is filled if its Geometry _n_ .NoFill cell is zero (0). If you indicate to Visio to fill an open **Path** object, it responds as if there is a LineTo cell from the **Path** object's endpoint to its begin point. When filling a **Path** object, Visio considers a point to be inside the **Path** object if a ray drawn from the point in any direction crosses the **Path** object or any of the shape's other **Path** objects cross an odd number of times.



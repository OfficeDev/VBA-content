---
title: Axis.CrossesAt Property (PowerPoint)
keywords: vbapp10.chm682006
f1_keywords:
- vbapp10.chm682006
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.CrossesAt
ms.assetid: ccc6438d-fb72-7674-0994-bf5d9cecf58d
ms.date: 06/08/2017
---


# Axis.CrossesAt Property (PowerPoint)

Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double**.


## Syntax

 _expression_. **CrossesAt**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting this property causes the  **[Crosses](axis-crosses-property-powerpoint.md)** property to change to **xlAxisCrossesCustom**.

You cannot use this property on radar charts. For 3-D charts, this property indicates where the plane defined by the category axes crosses the value axis.


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)


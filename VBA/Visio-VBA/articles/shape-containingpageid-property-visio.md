---
title: Shape.ContainingPageID Property (Visio)
keywords: vis_sdr.chm11260135
f1_keywords:
- vis_sdr.chm11260135
ms.prod: visio
api_name:
- Visio.Shape.ContainingPageID
ms.assetid: fd33d0d6-571d-47b5-28a7-6fa4aa671312
ms.date: 06/08/2017
---


# Shape.ContainingPageID Property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingPageID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.



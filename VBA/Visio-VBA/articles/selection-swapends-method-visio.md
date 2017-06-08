---
title: Selection.SwapEnds Method (Visio)
keywords: vis_sdr.chm11150895
f1_keywords:
- vis_sdr.chm11150895
ms.prod: visio
api_name:
- Visio.Selection.SwapEnds
ms.assetid: 515580db-4018-30b3-0ed6-cb3a412b62c7
ms.date: 06/08/2017
---


# Selection.SwapEnds Method (Visio)

Swaps the begin and endpoints of a one-dimensional (1-D) shape.


## Syntax

 _expression_ . **SwapEnds**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

The type of glue associated with the endpoints is also swapped. For example, if the begin point of a 1-D shape is glued to object A and the endpoint of the 1-D shape is not glued, after invoking the  **SwapEnds** method, the endpoint is glued to object A and the begin point is not glued.



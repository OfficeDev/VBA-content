---
title: Selection.ContainingPageID Property (Visio)
keywords: vis_sdr.chm11151930
f1_keywords:
- vis_sdr.chm11151930
ms.prod: visio
api_name:
- Visio.Selection.ContainingPageID
ms.assetid: f7d19685-9e1d-8867-978a-563dd3e93b0b
ms.date: 06/08/2017
---


# Selection.ContainingPageID Property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingPageID**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.



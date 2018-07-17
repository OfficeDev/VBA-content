---
title: Row.ContainingPageID Property (Visio)
keywords: vis_sdr.chm15851695
f1_keywords:
- vis_sdr.chm15851695
ms.prod: visio
api_name:
- Visio.Row.ContainingPageID
ms.assetid: 28a8e54d-fb2c-e6b6-ab18-ec71dc06eca5
ms.date: 06/08/2017
---


# Row.ContainingPageID Property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingPageID**

 _expression_ A variable that represents a **Row** object.


### Return Value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.



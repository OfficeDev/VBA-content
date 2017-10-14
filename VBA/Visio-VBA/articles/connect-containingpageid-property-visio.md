---
title: Connect.ContainingPageID Property (Visio)
keywords: vis_sdr.chm10351930
f1_keywords:
- vis_sdr.chm10351930
ms.prod: visio
api_name:
- Visio.Connect.ContainingPageID
ms.assetid: 4503f9e3-74ca-5948-ddc2-a91116faa588
ms.date: 06/08/2017
---


# Connect.ContainingPageID Property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingPageID**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.



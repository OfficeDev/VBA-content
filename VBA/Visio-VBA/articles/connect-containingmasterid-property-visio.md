---
title: Connect.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm10351935
f1_keywords:
- vis_sdr.chm10351935
ms.prod: visio
api_name:
- Visio.Connect.ContainingMasterID
ms.assetid: 4ac0f6c4-c5df-33e3-8c28-9bdf5d77d300
ms.date: 06/08/2017
---


# Connect.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Connect** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



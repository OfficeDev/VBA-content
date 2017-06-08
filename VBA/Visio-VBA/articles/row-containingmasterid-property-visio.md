---
title: Row.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm15851700
f1_keywords:
- vis_sdr.chm15851700
ms.prod: visio
api_name:
- Visio.Row.ContainingMasterID
ms.assetid: 12832d29-2eaf-ce37-fb30-ce2de24b140c
ms.date: 06/08/2017
---


# Row.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Row** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



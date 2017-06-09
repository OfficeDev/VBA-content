---
title: Cell.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm10151700
f1_keywords:
- vis_sdr.chm10151700
ms.prod: visio
api_name:
- Visio.Cell.ContainingMasterID
ms.assetid: 1daba8ed-69cd-2c80-8534-ba9fc4956292
ms.date: 06/08/2017
---


# Cell.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



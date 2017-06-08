---
title: Selection.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm11151935
f1_keywords:
- vis_sdr.chm11151935
ms.prod: visio
api_name:
- Visio.Selection.ContainingMasterID
ms.assetid: 9f9aad28-3e77-8ef8-29dc-e53852adf63d
ms.date: 06/08/2017
---


# Selection.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



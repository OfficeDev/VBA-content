---
title: Section.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm15751700
f1_keywords:
- vis_sdr.chm15751700
ms.prod: visio
api_name:
- Visio.Section.ContainingMasterID
ms.assetid: 642bc274-4977-8c1c-160f-b72c11bfbb1b
ms.date: 06/08/2017
---


# Section.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Section** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



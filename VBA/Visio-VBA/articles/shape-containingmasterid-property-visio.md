---
title: Shape.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm11260130
f1_keywords:
- vis_sdr.chm11260130
ms.prod: visio
api_name:
- Visio.Shape.ContainingMasterID
ms.assetid: e194cd7c-d7c0-2c08-a0df-764398efa447
ms.date: 06/08/2017
---


# Shape.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



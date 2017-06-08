---
title: Characters.ContainingMasterID Property (Visio)
keywords: vis_sdr.chm10251935
f1_keywords:
- vis_sdr.chm10251935
ms.prod: visio
api_name:
- Visio.Characters.ContainingMasterID
ms.assetid: 50ed7758-208e-15f0-14ac-801db910dabd
ms.date: 06/08/2017
---


# Characters.ContainingMasterID Property (Visio)

Returns the ID of the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMasterID**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Long


## Remarks

If the object is not in a  **Master** object, the **ContainingMasterID** property returns -1. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMasterID** property returns -1.



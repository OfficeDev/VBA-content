---
title: Shape.Master Property (Visio)
keywords: vis_sdr.chm11213870
f1_keywords:
- vis_sdr.chm11213870
ms.prod: visio
api_name:
- Visio.Shape.Master
ms.assetid: 698e205b-3cfc-2ee1-4fa1-73bc3d018b78
ms.date: 06/08/2017
---


# Shape.Master Property (Visio)

Returns the master from which the  **Shape** object was created. Read-only.


## Syntax

 _expression_ . **Master**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Master


## Remarks

If the  **Shape** object is not an instance of a master, its **Master** property returns **Nothing** .

If the  **Shape** object is in a group, its **Master** property is the same as the group's **Master** property.



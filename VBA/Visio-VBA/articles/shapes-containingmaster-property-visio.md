---
title: Shapes.ContainingMaster Property (Visio)
keywords: vis_sdr.chm11313300
f1_keywords:
- vis_sdr.chm11313300
ms.prod: visio
api_name:
- Visio.Shapes.ContainingMaster
ms.assetid: e7758236-92af-1a3a-fe1b-bce94a186eb9
ms.date: 06/08/2017
---


# Shapes.ContainingMaster Property (Visio)

Returns the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMaster**

 _expression_ A variable that represents a **Shapes** object.


### Return Value

Master


## Remarks

If the object isn't in a  **Master** object, the **ContainingMaster** property returns **Nothing** . For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMaster** property returns **Nothing** .



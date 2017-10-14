---
title: Shape.ContainingMaster Property (Visio)
keywords: vis_sdr.chm11213300
f1_keywords:
- vis_sdr.chm11213300
ms.prod: visio
api_name:
- Visio.Shape.ContainingMaster
ms.assetid: ca262f68-472e-3412-f620-ca837c40378c
ms.date: 06/08/2017
---


# Shape.ContainingMaster Property (Visio)

Returns the  **Master** object that contains an object. Read-only.


## Syntax

 _expression_ . **ContainingMaster**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Master


## Remarks

If the object isn't in a  **Master** object, the **ContainingMaster** property returns **Nothing** . For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMaster** property returns **Nothing** .



---
title: Shape.ContainingPage Property (Visio)
keywords: vis_sdr.chm11213305
f1_keywords:
- vis_sdr.chm11213305
ms.prod: visio
api_name:
- Visio.Shape.ContainingPage
ms.assetid: 18fe6146-34eb-9369-603b-b3b316aa23d7
ms.date: 06/08/2017
---


# Shape.ContainingPage Property (Visio)

Returns the page that contains an object.


## Syntax

 _expression_ . **ContainingPage**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Page


## Remarks

If the object isn't in a  **Page** object, the **ContainingPage** property returns **Nothing** . For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPage** property returns **Nothing** .



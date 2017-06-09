---
title: Shapes.ContainingPage Property (Visio)
keywords: vis_sdr.chm11313305
f1_keywords:
- vis_sdr.chm11313305
ms.prod: visio
api_name:
- Visio.Shapes.ContainingPage
ms.assetid: 0e74569b-7044-6743-9dfe-52ff8acb11dc
ms.date: 06/08/2017
---


# Shapes.ContainingPage Property (Visio)

Returns the page that contains an object.


## Syntax

 _expression_ . **ContainingPage**

 _expression_ A variable that represents a **Shapes** object.


### Return Value

Page


## Remarks

If the object isn't in a  **Page** object, the **ContainingPage** property returns **Nothing** . For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPage** property returns **Nothing** .



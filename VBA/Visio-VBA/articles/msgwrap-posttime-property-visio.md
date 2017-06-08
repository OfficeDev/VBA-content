---
title: MSGWrap.posttime Property (Visio)
keywords: vis_sdr.chm16150770
f1_keywords:
- vis_sdr.chm16150770
ms.prod: visio
api_name:
- Visio.MSGWrap.posttime
ms.assetid: e43c865b-eca8-22c7-de8e-1c6ec3f53348
ms.date: 06/08/2017
---


# MSGWrap.posttime Property (Visio)

Gets or sets the  **time** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_ . **posttime**

 _expression_ A variable that represents a **MSGWrap** object.


### Return Value

Long


## Remarks

The  **posttime** property corresponds to the **time** member of the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.



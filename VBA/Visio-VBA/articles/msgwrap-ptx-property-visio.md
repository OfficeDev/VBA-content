---
title: MSGWrap.ptx Property (Visio)
keywords: vis_sdr.chm16150810
f1_keywords:
- vis_sdr.chm16150810
ms.prod: visio
api_name:
- Visio.MSGWrap.ptx
ms.assetid: 2dfae0e3-c78c-7beb-9edc-5b269d7f7c33
ms.date: 06/08/2017
---


# MSGWrap.ptx Property (Visio)

Gets or sets the  **pt.x** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_ . **ptx**

 _expression_ A variable that represents a **MSGWrap** object.


### Return Value

Long


## Remarks

The  **ptx** property corresponds to the **pt.x** member in the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.



---
title: MSGWrap.wParam Property (Visio)
keywords: vis_sdr.chm16150930
f1_keywords:
- vis_sdr.chm16150930
ms.prod: visio
api_name:
- Visio.MSGWrap.wParam
ms.assetid: 1f0e1aa9-63ea-e73e-2e51-8eb3e4bd8393
ms.date: 06/08/2017
---


# MSGWrap.wParam Property (Visio)

Gets or sets the  **wParam** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_ . **wParam**

 _expression_ A variable that represents a **MSGWrap** object.


### Return Value

Long


## Remarks

The  **wParam** property corresponds to the **wParam** member ofn the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.



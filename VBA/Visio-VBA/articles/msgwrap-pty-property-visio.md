---
title: MSGWrap.pty Property (Visio)
keywords: vis_sdr.chm16150805
f1_keywords:
- vis_sdr.chm16150805
ms.prod: visio
api_name:
- Visio.MSGWrap.pty
ms.assetid: 1f3ecb54-db9c-4d07-fbe2-a63b2e4be083
ms.date: 06/08/2017
---


# MSGWrap.pty Property (Visio)

Gets or sets the  **pt.y** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_ . **pty**

 _expression_ A variable that represents a **MSGWrap** object.


### Return Value

Long


## Remarks

The  **pty** property corresponds to the **pt.y** member in the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.



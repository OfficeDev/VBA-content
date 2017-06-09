---
title: MSGWrap.lParam Property (Visio)
keywords: vis_sdr.chm16150705
f1_keywords:
- vis_sdr.chm16150705
ms.prod: visio
api_name:
- Visio.MSGWrap.lParam
ms.assetid: 52995a4b-f139-6ecc-a5fe-23882a3f2245
ms.date: 06/08/2017
---


# MSGWrap.lParam Property (Visio)

Gets or sets the  **lParam** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_ . **lParam**

 _expression_ A variable that represents a **MSGWrap** object.


### Return Value

Long


## Remarks

The  **lParam** property corresponds to the **lParam** member of the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.



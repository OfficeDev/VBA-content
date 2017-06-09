---
title: Window.ScrollLock Property (Visio)
keywords: vis_sdr.chm11651650
f1_keywords:
- vis_sdr.chm11651650
ms.prod: visio
api_name:
- Visio.Window.ScrollLock
ms.assetid: 08459237-ff58-cd39-319f-60d7bb730487
ms.date: 06/08/2017
---


# Window.ScrollLock Property (Visio)

Determines whether scrolling is disabled in a Microsoft Visio window. Read/write.


## Syntax

 _expression_ . **ScrollLock**

 _expression_ A variable that represents a **Window** object.


### Return Value

Boolean


## Remarks

Scrolling ( **False** ) is the default Visio behavior. You can set the **ScrollLock** property to **True** to prevent scrolling in any Visio window, including docked stencil windows, but not including anchored windows.

The  **ScrollLock** property setting is valid only for a given window at run time, and is not persisted (saved) in either the Visio document or the registry.



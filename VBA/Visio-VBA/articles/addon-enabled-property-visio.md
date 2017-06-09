---
title: Addon.Enabled Property (Visio)
keywords: vis_sdr.chm12413455
f1_keywords:
- vis_sdr.chm12413455
ms.prod: visio
api_name:
- Visio.Addon.Enabled
ms.assetid: fcc719d3-7b1c-e356-6f92-7717ecea10df
ms.date: 06/08/2017
---


# Addon.Enabled Property (Visio)

Determines whether or not an  **Addon** object is currently enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents an **Addon** object.


### Return Value

Integer


## Remarks

An add-on implemented by an executable (EXE) file always reports itself as enabled. An add-on implemented by a Visio Solutions Library (VSL) file reports itself as enabled or disabled according to the enabling policy that the VSL file has registered for that add-on.

You cannot tell an add-on to enable or disable itself. Visio will not send a run message to a disabled add-on. If an add-on is disabled, its name appears unavailable in the Visio user interface.



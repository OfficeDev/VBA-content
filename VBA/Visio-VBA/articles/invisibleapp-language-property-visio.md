---
title: InvisibleApp.Language Property (Visio)
keywords: vis_sdr.chm17513800
f1_keywords:
- vis_sdr.chm17513800
ms.prod: visio
api_name:
- Visio.InvisibleApp.Language
ms.assetid: e8f7408a-5589-d4b4-0e85-95ac714f7e6f
ms.date: 06/08/2017
---


# InvisibleApp.Language Property (Visio)

Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.


## Syntax

 _expression_ . **Language**( **_lpi4Ret_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Long


## Remarks

The  **Language** property returns the language ID recorded in the object's VERSIONINFO resource. The IDs returned are the standard IDs used by Microsoft Windows to encode different language versions. For example, the **Language** property returns &;H0409 for the U.S. English version of Visio. For details, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.



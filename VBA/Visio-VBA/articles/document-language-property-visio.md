---
title: Document.Language Property (Visio)
keywords: vis_sdr.chm10551705
f1_keywords:
- vis_sdr.chm10551705
ms.prod: visio
api_name:
- Visio.Document.Language
ms.assetid: 76f995fd-8b4d-7292-50c1-8dcb6448c2ec
ms.date: 06/08/2017
---


# Document.Language Property (Visio)

Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.


## Syntax

 _expression_ . **Language**

 _expression_ A variable that represents a **Document** object.


### Return Value

Long


## Remarks

The  **Language** property returns the language ID recorded in the object's VERSIONINFO resource. The IDs returned are the standard IDs used by Microsoft Windows to encode different language versions. For example, the **Language** property returns &;H0409 for the U.S. English version of Visio. For details, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.



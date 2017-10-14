---
title: Shape.Language Property (Visio)
keywords: vis_sdr.chm11251705
f1_keywords:
- vis_sdr.chm11251705
ms.prod: visio
api_name:
- Visio.Shape.Language
ms.assetid: 6c7ab4ca-8813-9cbc-d433-a3991a0b450f
ms.date: 06/08/2017
---


# Shape.Language Property (Visio)

Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.


## Syntax

 _expression_ . **Language**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

The  **Language** property returns the language ID recorded in the object's VERSIONINFO resource. The IDs returned are the standard IDs used by Microsoft Windows to encode different language versions. For example, the **Language** property returns &;H0409 for the U.S. English version of Visio. For details, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.



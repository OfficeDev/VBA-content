---
title: Document.Path Property (Visio)
keywords: vis_sdr.chm10514050
f1_keywords:
- vis_sdr.chm10514050
ms.prod: visio
api_name:
- Visio.Document.Path
ms.assetid: 50c20d69-3909-9383-1d2c-d1744a96e751
ms.date: 06/08/2017
---


# Document.Path Property (Visio)

Returns the drive and folder path of the Microsoft Visio document. Read-only.


## Syntax

 _expression_ . **Path**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

If the document has not been saved, the  **Path** property of the **Document** object returns a zero-length string ("").



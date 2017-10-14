---
title: Document.FullName Property (Visio)
keywords: vis_sdr.chm10513595
f1_keywords:
- vis_sdr.chm10513595
ms.prod: visio
api_name:
- Visio.Document.FullName
ms.assetid: 9f6d15ab-9913-57f4-a0ee-57618d5b1b0f
ms.date: 06/08/2017
---


# Document.FullName Property (Visio)

Returns the name of a document, including the drive and path. Read-only.


## Syntax

 _expression_ . **FullName**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

Use the  **FullName** property to obtain a document's drive, folder path, and file name as one string. The returned value can include UNC drive names (for example, \\ _drive\folder_ ).



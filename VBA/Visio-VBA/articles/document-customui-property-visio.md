---
title: Document.CustomUI Property (Visio)
keywords: vis_sdr.chm10562690
f1_keywords:
- vis_sdr.chm10562690
ms.prod: visio
api_name:
- Visio.Document.CustomUI
ms.assetid: dff5841d-f2cc-c8fd-1b30-ca0145f5c04c
ms.date: 06/08/2017
---


# Document.CustomUI Property (Visio)

Gets or sets the Ribbon XML string that is passed to the document to customize the Microsoft Office Fluent user interface. Read/write.


## Syntax

 _expression_ . **CustomUI**

 _expression_ A variable that represents a **[Document](document-object-visio.md)** object.


### Return Value

 **String**


## Remarks

When you set the  **CustomUI** property value, Microsoft Visio does not perform validation on the Ribbon XML. Instead, the XML is persisted in the document file and validated the next time that the document loads.



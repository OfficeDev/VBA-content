---
title: Document.UserCustomUI Property (Visio)
keywords: vis_sdr.chm10562695
f1_keywords:
- vis_sdr.chm10562695
ms.prod: visio
api_name:
- Visio.Document.UserCustomUI
ms.assetid: cdd28d78-a75a-b8c4-71e9-74c24ee9ecf1
ms.date: 06/08/2017
---


# Document.UserCustomUI Property (Visio)

Gets or sets the Ribbon XML string that is passed to the document to customize the  **Quick Access** toolbar or the Ribbon. Read/write.


## Syntax

 _expression_ . **UserCustomUI**

 _expression_ A variable that represents a **[Document](document-object-visio.md)** object.


## Remarks

When you set the  **UserCustomUI** property value, Microsoft Visio does not validate the Ribbon XML. Instead, the XML is persisted in the document file and validated the next time that the document is loaded.



---
title: Document.ContainsWorkspaceEx Property (Visio)
keywords: vis_sdr.chm10550530
f1_keywords:
- vis_sdr.chm10550530
ms.prod: visio
api_name:
- Visio.Document.ContainsWorkspaceEx
ms.assetid: 074d4b49-cb26-5d11-d710-7d27f2f4dd01
ms.date: 06/08/2017
---


# Document.ContainsWorkspaceEx Property (Visio)

Gets or sets whether workspace information is saved with the document. Read/write.


## Syntax

 _expression_ . **ContainsWorkspaceEx**

 _expression_ An expression that returns a **Document** object.


### Return Value

Boolean


## Remarks

The  **ContainsWorkspaceEx** property setting corresponds to the setting of the **Save workspace** check box on the **Summary** tab of the **Document Properties** dialog box in the Microsoft Visio user interface.

The  **ContainsWorkspaceEx** property replaces the read-only **ContainsWorkspace** property that existed in previous versions of Visio. Note that although the **ContainsWorkspace** property is now deprecated, it still returns a value.

Note also that in previous versions of Visio, to save a document along with workspace information, you used the  **[Document.SaveAs](document-saveas-method-visio.md)** method, passing it the **visSaveAsWS** constant.

Starting with Visio 2007, by default, workspace information is saved with all documents except stencils. To specify that workspace information not be saved with a document, set  **ContainsWorkspaceEx** to **False** .



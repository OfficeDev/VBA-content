---
title: Document.Close Method (Visio)
keywords: vis_sdr.chm10516125
f1_keywords:
- vis_sdr.chm10516125
ms.prod: visio
api_name:
- Visio.Document.Close
ms.assetid: 913572fd-cacb-8d06-0e5f-3bd2e98d6d13
ms.date: 06/08/2017
---


# Document.Close Method (Visio)

Closes a document.


## Syntax

 _expression_ . **Close**

 _expression_ A variable that represents a **Document** object.


### Return Value

Nothing


## Remarks

If the indicated window is the only window open for a document and the document contains unsaved changes, an alert appears asking if you want to save the document. You can use the  **AlertResponse** property to prevent the alert from appearing.

If you close a docked stencil window, only that window is closed. However, if you close a drawing window that contains docked stencils, the docked stencil window is also closed.



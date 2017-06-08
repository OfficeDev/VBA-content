---
title: CoauthMergeEvent.WorkingDocument Property (Visio)
ms.prod: visio
ms.assetid: 0f3c4358-0d63-df7f-12fe-7f378bacca86
ms.date: 06/08/2017
---


# CoauthMergeEvent.WorkingDocument Property (Visio)

Returns a [Document](document-object-visio.md) object that represents a merged document that includes changes by the current user only. Read-only.


## Syntax

 _expression_ . **WorkingDocument**

 _expression_ A variable that represents a **CoauthMergeEvent** object.


## Remarks

Changes to the merged document returned by the  **WorkingDocument** property are what fire the[Document.AfterDocumentMerge](document-afterdocumentmerge-event-visio.md) or[Documents.AfterDcoumentMerge](documents-afterdocumentmerge-event-visio.md) event represented by the specified **CoauthMergeEvent** object.


## Property value

 **IVDOCUMENT**


## See also


#### Other resources


[CoauthMergeEvent Object](coauthmergeevent-object-visio.md)



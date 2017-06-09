---
title: Document.DocumentSavedAs Event (Visio)
keywords: vis_sdr.chm10519140
f1_keywords:
- vis_sdr.chm10519140
ms.prod: visio
api_name:
- Visio.Document.DocumentSavedAs
ms.assetid: 36714188-964b-880b-9504-62a6a50179f1
ms.date: 06/08/2017
---


# Document.DocumentSavedAs Event (Visio)

Occurs after a document is saved by using the  **Save As** command.


## Syntax

Private Sub  _expression_ _**DocumentSavedAs**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that was saved.|

## Remarks

The  **DocumentSavedAs** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

If the  **DocumentSavedAs** event is fired because a save was initiated by a user or a program, the **EventInfo** property returns the following string:

 "/saveasfile=<filename>"

If it fires because Microsoft Visio is saving a copy of an open file (for autorecovery or to include as a mail attachment), the  **EventInfo** property returns one of the following strings:




- If the event is fired for autorecovery purposes, the name of a recovery file in this format: "/autosavefile= _drivename:\foldername\filename_ "
    
- If the event is fired because a document copy is being made to send as a mail attachment, the name of an attachment file in this format: "/mailfile= _drivename:\foldername\filename_ "
    


If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).

If you are handling this event from a program that receives a notification over a connection created by using the  **AddAdvise** method, the _varMoreInfo_ argument to **VisEventProc** designates the document index: "/doc=1".



---
title: Documents.BeforeDocumentSaveAs Event (Visio)
keywords: vis_sdr.chm10619035
f1_keywords:
- vis_sdr.chm10619035
ms.prod: visio
api_name:
- Visio.Documents.BeforeDocumentSaveAs
ms.assetid: f5ee2f06-9559-470b-bf32-23dad5a5cc0a
ms.date: 06/08/2017
---


# Documents.BeforeDocumentSaveAs Event (Visio)

Occurs just before a document is saved by using the  **Save As** command.


## Syntax

Private Sub  _expression_ _**BeforeDocumentSaveAs**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be saved.|

## Remarks

The  **BeforeDocumentSaveAs** event fires when a document is saved to either a native format (for example, VSD or VDX) or a non-native format (for example, HTM or BMP). It does not fire when a document is saved to DWG, DXF, and DGN formats. To save a document in a non-native format programmatically, you must use the **Export** method of the **Page** object. Note that when you call the **SaveAs** method, Microsoft Visio fires first the **BeforeDocumentSaveAs** event and then the **DocumentSavedAs** event. Calling the **Export** method, however, fires the **BeforeDocumentSaveAs** event but not the **DocumentSavedAs** event that follows it in response to the **SaveAs** method.

The  **BeforeDocumentSaveAs** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

If the  **BeforeDocumentSaveAs** event is fired because a save was initiated by a user or a program, the **EventInfo** property returns the following string:

"/saveasfile=<filename>"

If it fires because Visio is saving a copy of an open file (for autorecovery or to include as a mail attachment), the  **EventInfo** property will return one of the following strings:


- If the event is fired for autorecovery purposes, the name of a recovery file in this format: "/autosavefile=C:\TEMP\~$2VSO2FD.vsd"
    
- If the event is fired because a document copy is being made to send as a mail attachment, the name of an attachment file in this format: "/mailfile=C:\TEMP\~$2VSO2FD.vsd"
    
If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).

If you are handling this event from a program that receives a notification over a connection by using the  **AddAdvise** method, the _vMoreInfo_ argument to **VisEventProc** designates the document index: "/doc=1".



---
title: Document.Sync Event (Word)
keywords: vbawd10.chm4001007
f1_keywords:
- vbawd10.chm4001007
ms.prod: word
api_name:
- Word.Document.Sync
ms.assetid: cc46cfdf-ae26-9bba-7084-64349859d304
ms.date: 06/08/2017
---


# Document.Sync Event (Word)

This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## Syntax

Private Sub  _expression_ _**Sync**( **_SyncEventType_** )

 _expression_ A variable that represents a **[Document](document-object-word.md)** object that has been declared using the **WithEvents** keyword in a class module. For information about using events with the **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SyncEventType_|Required| **MsoSyncEventType**|The status of the document synchronization.|

## Example

The following example displays a message if the synchronization of a document in a Document Workspace fails.


```vb
Private Sub Document_Sync(ByVal SyncEventType As Office.MsoSyncEventType) 
 
 If SyncEventType = msoSyncEventDownloadFailed Or _ 
 SyncEventType = msoSyncEventUploadFailed Then 
 
 MsgBox "Document synchronization failed. " &; _ 
 "Please contact your administrator " &; vbCrLf &; _ 
 "or try again later." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)


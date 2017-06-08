---
title: Application.DocumentSync Event (Word)
keywords: vbawd10.chm4000027
f1_keywords:
- vbawd10.chm4000027
ms.prod: word
api_name:
- Word.Application.DocumentSync
ms.assetid: 9c83f692-8d05-2c52-11ef-46ac0ff69431
ms.date: 06/08/2017
---


# Application.DocumentSync Event (Word)

This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## Syntax

Private Sub_**DocumentSync**( **_Doc_** , **_SyncEventType_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object declared using the **WithEvents** keyword in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document being synchronized.|
| _SyncEventType_|Required| **MsoSyncEventType**|The status of the document synchronization.|

## Remarks

For information about using events with the  **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

The following example displays a message if the synchronization of a document in a Document Workspace fails.


```vb
Private Sub app_DocumentSync(ByVal Doc As Document, _ 
 ByVal SyncEventType As Office.MsoSyncEventType) 
 
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


[Application Object](application-object-word.md)


---
title: Document.Sync Property (Word)
keywords: vbawd10.chm158007762
f1_keywords:
- vbawd10.chm158007762
ms.prod: word
api_name:
- Word.Document.Sync
ms.assetid: c48b0b07-84c6-0097-509c-ee6fb9b3784e
ms.date: 06/08/2017
---


# Document.Sync Property (Word)

This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## Syntax

 _expression_ . **Sync**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

The following example displays the name of the last person to modify the active document if the active document is a shared document in a Document Workspace.


```vb
Dim eStatus As MsoSyncStatusType 
Dim strLastUser As String 
 
eStatus = ActiveDocument.Sync.Status 
 
If eStatus = msoSyncStatusLatest Then 
 strLastUser = ActiveDocument.Sync.WorkspaceLastChangedBy 
 MsgBox "You have the most up-to-date copy." &; _ 
 "This file was last modified by " &; strLastUser 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)


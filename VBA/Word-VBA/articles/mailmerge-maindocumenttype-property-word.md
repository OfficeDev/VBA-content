---
title: MailMerge.MainDocumentType Property (Word)
keywords: vbawd10.chm153092097
f1_keywords:
- vbawd10.chm153092097
ms.prod: word
api_name:
- Word.MailMerge.MainDocumentType
ms.assetid: 6275d472-b513-1879-e48a-326f21d6321d
ms.date: 06/08/2017
---


# MailMerge.MainDocumentType Property (Word)

Returns or sets the mail merge main document type. Read/write  **WdMailMergeMainDocType** .


## Syntax

 _expression_ . **MainDocumentType**

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example creates a new document and makes it a catalog main document for a mail merge operation.


```vb
Set myDoc = Documents.Add 
myDoc.MailMerge.MainDocumentType = wdCatalog
```

This example determines whether the active document is a main document for a mail merge operation, and then it displays a message in the status bar.




```vb
Set doc = ActiveDocument 
If doc.MailMerge.MainDocumentType = wdNotAMergeDocument Then 
 StatusBar = "Not a mail merge main document" 
Else 
 StatusBar = "Document is a mail merge main document." 
End If
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)


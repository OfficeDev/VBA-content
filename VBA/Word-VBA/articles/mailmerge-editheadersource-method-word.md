---
title: MailMerge.EditHeaderSource Method (Word)
keywords: vbawd10.chm153092204
f1_keywords:
- vbawd10.chm153092204
ms.prod: word
api_name:
- Word.MailMerge.EditHeaderSource
ms.assetid: d1be3c68-b7f8-7591-2a1a-b5898f731fc6
ms.date: 06/08/2017
---


# MailMerge.EditHeaderSource Method (Word)

Opens the header source attached to a mail merge main document, or activates the header source if it is already open.


## Syntax

 _expression_ . **EditHeaderSource**

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Remarks

If the mail merge main document doesn't have a header source, this method causes an error.


## Example

This example attaches a header source to the active document and then opens the header source.


```vb
With ActiveDocument.MailMerge 
 .MainDocumentType = wdFormLetters 
 .OpenHeaderSource Name:="C:\Documents\Header.doc" 
 .EditHeaderSource 
End With
```

This example opens the header source if the active document has an associated header file attached to it.




```vb
Dim mmTemp As MailMerge 
 
Set mmTemp = ActiveDocument.MailMerge 
If mmTemp.State = wdMainAndSourceAndHeader Or _ 
 mmTemp.State = wdMainAndHeader Then 
 mmTemp.EditHeaderSource 
End If
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)


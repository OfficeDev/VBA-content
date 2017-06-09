---
title: MailMerge.SuppressBlankLines Property (Word)
keywords: vbawd10.chm153092103
f1_keywords:
- vbawd10.chm153092103
ms.prod: word
api_name:
- Word.MailMerge.SuppressBlankLines
ms.assetid: 27faf7f7-5d7b-2377-0775-80ce6d13eb64
ms.date: 06/08/2017
---


# MailMerge.SuppressBlankLines Property (Word)

 **True** if blank lines are suppressed when mail merge fields in a mail merge main document are empty. Read/write **Boolean** .


## Syntax

 _expression_ . **SuppressBlankLines**

 _expression_ An expression that returns a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example opens Main.doc and executes the mail merge operation. When merge fields are empty, blank lines are suppressed in the merge document.


```vb
Set myDoc = Documents.Open(FileName:="C:\My Documents\Main.doc") 
With myDoc.MailMerge 
 .SuppressBlankLines = True 
 .Destination = wdSendToPrinter 
 .Execute 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)


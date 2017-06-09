---
title: Document.MailMerge Property (Word)
keywords: vbawd10.chm158007323
f1_keywords:
- vbawd10.chm158007323
ms.prod: word
api_name:
- Word.Document.MailMerge
ms.assetid: 71c144ab-b1fb-c031-2e8d-54e9802fab5d
ms.date: 06/08/2017
---


# Document.MailMerge Property (Word)

Returns a  **[MailMerge](mailmerge-object-word.md)** object that represents the mail merge functionality for the specified document. Read-only.


## Syntax

 _expression_ . **MailMerge**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **MailMerge** object is available regardless of whether the specified document is a mail merge main document. Use the **State** property to determine the current state of the mail merge operation.


## Example

This example executes a mail merge if the active document is a main document with an attached data source.


```vb
Set myMerge = ActiveDocument.MailMerge 
If myMerge.State = wdMainAndDataSource Then myMerge.Execute
```

This example merges the main document with records 1 through 4 and sends the merge documents to the printer.




```vb
With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = 1 
 .DataSource.LastRecord = 4 
 .Destination = wdSendToPrinter 
 .SuppressBlankLines = True 
 .Execute 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)


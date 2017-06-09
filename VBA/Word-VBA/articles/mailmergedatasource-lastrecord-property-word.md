---
title: MailMergeDataSource.LastRecord Property (Word)
keywords: vbawd10.chm152895497
f1_keywords:
- vbawd10.chm152895497
ms.prod: word
api_name:
- Word.MailMergeDataSource.LastRecord
ms.assetid: 9c51a46f-5d46-c066-5cc5-6bcd0a124209
ms.date: 06/08/2017
---


# MailMergeDataSource.LastRecord Property (Word)

Returns or sets the number of the last record to be merged in a mail merge operation. Read/write  **Long** .


## Syntax

 _expression_ . **LastRecord**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Example

This example merges the main document with records 2 through 4 and sends the merge documents to a new document.


```vb
With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = 2 
 .DataSource.LastRecord = 4 
 .Destination = wdSendToNewDocument 
 .Execute 
End With
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


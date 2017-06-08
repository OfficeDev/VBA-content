---
title: MailMergeDataSource.FirstRecord Property (Word)
keywords: vbawd10.chm152895496
f1_keywords:
- vbawd10.chm152895496
ms.prod: word
api_name:
- Word.MailMergeDataSource.FirstRecord
ms.assetid: c94e1581-a6eb-84e0-6acc-f8ca6ae7575b
ms.date: 06/08/2017
---


# MailMergeDataSource.FirstRecord Property (Word)

Returns or sets the number of the first record to be merged in a mail merge operation. Read/write  **Long** .


## Syntax

 _expression_ . **FirstRecord**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Example

This example merges the main document with records 1 through 3 and sends the merge documents to the printer.


```vb
With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = 1 
 .DataSource.LastRecord = 3 
 .Destination = wdSendToPrinter 
 .Execute 
End With
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


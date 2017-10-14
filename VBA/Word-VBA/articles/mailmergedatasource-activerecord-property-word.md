---
title: MailMergeDataSource.ActiveRecord Property (Word)
keywords: vbawd10.chm152895495
f1_keywords:
- vbawd10.chm152895495
ms.prod: word
api_name:
- Word.MailMergeDataSource.ActiveRecord
ms.assetid: bbac1bf4-e11a-887c-8502-0bc95c57bcf9
ms.date: 06/08/2017
---


# MailMergeDataSource.ActiveRecord Property (Word)

Returns or sets the active mail merge record. Can be either a valid record number in the query result or one of the  **WdMailMergeActiveRecord** constants.


## Syntax

 _expression_ . **ActiveRecord**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

The active record number is the position of the record in the query result produced by the current query options; as such, this number isn't necessarily the position of the record in the data source.


## Example

This example hides the mail merge field codes in the active document so that the merge data is visible in the main document. The active record is then advanced to the next record in the data source.


```vb
If ActiveDocument.MailMerge.MainDocumentType <> _ 
 wdNotAMergeDocument Then 
 With ActiveDocument.MailMerge 
 .ViewMailMergeFieldCodes = False 
 .DataSource.ActiveRecord = wdNextRecord 
 End With 
End If
```

This example returns the numeric position of the active record from Main2.doc.




```vb
Dim intRecordNumber as Integer 
 
If Documents("Main2.doc").MailMerge.State = _ 
 wdMainAndDataSource Or _ 
 wdMainAndSourceAndHeader Then 
 intRecordNumber = Documents("Main2.doc").MailMerge _ 
 .DataSource.ActiveRecord 
End If
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


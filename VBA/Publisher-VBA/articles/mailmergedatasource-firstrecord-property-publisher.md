---
title: MailMergeDataSource.FirstRecord Property (Publisher)
keywords: vbapb10.chm6291464
f1_keywords:
- vbapb10.chm6291464
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.FirstRecord
ms.assetid: e6eefea9-b353-27ff-d8e4-dc135c0c4665
ms.date: 06/08/2017
---


# MailMergeDataSource.FirstRecord Property (Publisher)

Returns or sets a  **Long** that represents the number of the first record to be merged in a mail merge or catalog merge operation. Read/write.


## Syntax

 _expression_. **FirstRecord**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Long


## Example

This example sets the active record as the first record to be merged, and then merges three records ending with the record two records forward in the data source. This example assumes that the active publication is a mail merge document.


```vb
Sub RecordOne() 
 With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = .DataSource.ActiveRecord 
 .DataSource.LastRecord = .DataSource.ActiveRecord + 2 
 .Execute Pause:=True 
 End With 
End Sub
```



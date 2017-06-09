---
title: MailMergeDataSource.LastRecord Property (Publisher)
keywords: vbapb10.chm6291474
f1_keywords:
- vbapb10.chm6291474
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.LastRecord
ms.assetid: c1d11d3e-5f6f-2729-081b-5727c75fbc8d
ms.date: 06/08/2017
---


# MailMergeDataSource.LastRecord Property (Publisher)

Returns or sets a  **Long** that represents the number of the last record to be merged in a mail merge or catalog merge operation. Read/write.


## Syntax

 _expression_. **LastRecord**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Long


## Example

This example sets the active record as the first record to be merged and then sets the last record as the record two records forward in the data source. This example assumes that the active publication is a mail merge publication.


```vb
Sub RecordOne() 
 With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = .DataSource.ActiveRecord 
 .DataSource.LastRecord = .DataSource.ActiveRecord + 2 
 .Execute Pause:=True 
 End With 
End Sub
```



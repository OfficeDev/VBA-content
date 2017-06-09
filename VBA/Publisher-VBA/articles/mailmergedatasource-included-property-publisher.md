---
title: MailMergeDataSource.Included Property (Publisher)
keywords: vbapb10.chm6291465
f1_keywords:
- vbapb10.chm6291465
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.Included
ms.assetid: 1cdac925-5fd6-e1d0-4612-0641e6057a7e
ms.date: 06/08/2017
---


# MailMergeDataSource.Included Property (Publisher)

 **True** if a record is included in a mail merge. Read/write **Boolean**.


## Syntax

 _expression_. **Included**

 _expression_A variable that represents an  **MailMergeDataSource** object.


### Return Value

Boolean


## Remarks

Use the  **[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-publisher.md)** method to set the included status for all mail merge records.


## Example

This example searches the records to verify that the length of the PostalCode field for each record is at least five digits long. If it is not, the record is excluded from the mail merge and flagged as invalid.


```vb
Sub ExcludeRecords() 
 Dim intRecord As Integer 
 With ActiveDocument.MailMerge 
 For intRecord = 1 To .DataSource.RecordCount 
 .DataSource.ActiveRecord = intRecord 
 If Len(.DataSource.DataFields("PostalCode").Value) < 5 Then 
 With .DataSource 
 .Included = False 
 .InvalidAddress = True 
 .InvalidComments = "This record is removed " &; _ 
 "from the mail merge because its postal code" &; _ 
 "has less than five digits." 
 End With 
 End If 
 Next 
 End With 
End Sub
```



---
title: MailMergeDataSource.InvalidComments Property (Publisher)
keywords: vbapb10.chm6291473
f1_keywords:
- vbapb10.chm6291473
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.InvalidComments
ms.assetid: ee08b03a-57e2-d79c-ee9f-a6f9231c8d6b
ms.date: 06/08/2017
---


# MailMergeDataSource.InvalidComments Property (Publisher)

If the  **[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)** property is **True**, this property returns or sets a  **String** that describes invalid data in a mail merge record. Read/write.


## Syntax

 _expression_. **InvalidComments**

 _expression_A variable that represents an  **MailMergeDataSource** object.


### Return Value

String


## Remarks

Use the  **[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-publisher.md)** method to set both the **[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)** and  **InvalidComments** properties for all records in a data source.


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



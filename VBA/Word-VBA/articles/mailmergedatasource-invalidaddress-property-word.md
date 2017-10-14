---
title: MailMergeDataSource.InvalidAddress Property (Word)
keywords: vbawd10.chm152895502
f1_keywords:
- vbawd10.chm152895502
ms.prod: word
api_name:
- Word.MailMergeDataSource.InvalidAddress
ms.assetid: ac84a87e-2125-851d-90ab-42359898edcc
ms.date: 06/08/2017
---


# MailMergeDataSource.InvalidAddress Property (Word)

 **True** for Microsoft Word to mark a record in a mail merge data source if it contains invalid data in an address field. Read/write **Boolean** .


## Syntax

 _expression_ . **InvalidAddress**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

Use the  **[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-word.md)** method to set both the **InvalidAddress** and **[InvalidComments](mailmergedatasource-invalidcomments-property-word.md)** properties for all records in a data source.


## Example

This example loops through the records in the mail merge data source and checks whether the ZIP Code field (in this case field number six) contains fewer than five digits. If a record does contain a ZIP Code of fewer than five digits, the record is excluded from the mail merge and the address is marked as invalid.


```vb
Sub ExcludeRecords() 
 
 Dim intCount As Integer 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 Do 
 intCount = intCount + 1 
 'Counts the number of digits in the postal code field and if 
 'it is less than 5, the record is excluded from the mail merge, 
 'marked as having an invalid address, and given a comment 
 'describing why the postal code was removed 
 If Len(.DataFields(6).Value) < 5 Then 
 .Included = False 
 .InvalidAddress = True 
 .InvalidComments = "The ZIP Code for this record" &; _ 
 "has fewer than five digits. It will be" &; _ 
 "removed from the mail merge process." 
 End If 
 
 .ActiveRecord = wdNextRecord 
 Loop Until intCount >= .ActiveRecord 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


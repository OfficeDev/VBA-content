---
title: MailMergeDataSource.InvalidComments Property (Word)
keywords: vbawd10.chm152895503
f1_keywords:
- vbawd10.chm152895503
ms.prod: word
api_name:
- Word.MailMergeDataSource.InvalidComments
ms.assetid: 4eb0ea4d-e89d-548d-f3be-1d0c3592ce53
ms.date: 06/08/2017
---


# MailMergeDataSource.InvalidComments Property (Word)

If the  **[InvalidAddress](mailmergedatasource-invalidaddress-property-word.md)** property is **True** , returns or sets a **String** that describes an invalid address error. Read/write.


## Syntax

 _expression_ . **InvalidComments**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

Use the  **[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-word.md)** method to set both the **InvalidAddress** and **InvalidComments** properties for all records in a data source.


## Example

This example loops through the records in the mail merge data source and checks whether the ZIP Code field (in this case field number six) contains fewer than five digits. If a record does contain a ZIP Code of fewer than five digits, the record is excluded from the mail merge, the address is marked as invalid, and a comment about why the record was excluded is added.


```vb
Sub ExcludeRecords() 
 
 Dim intCount As Integer 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 Do 
 intCount = intCount + 1 
 'Counts the number of digits in the postal code field and if 
 'it is fewer than 5, the record is excluded from the mail merge, 
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


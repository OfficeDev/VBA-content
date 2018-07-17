---
title: MailMergeDataSource.Included Property (Word)
keywords: vbawd10.chm152895501
f1_keywords:
- vbawd10.chm152895501
ms.prod: word
api_name:
- Word.MailMergeDataSource.Included
ms.assetid: 7d82056d-111c-27ce-a61c-be5876ee47df
ms.date: 06/08/2017
---


# MailMergeDataSource.Included Property (Word)

 **True** if a record is included in a mail merge. Read/write **Boolean** .


## Syntax

 _expression_ . **Included**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

Use the  **[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-word.md)** method to include or exclude all records in a mail merge data source.


## Example

This example loops through the records in the mail merge data source and checks if the ZIP Code field (in this case field number six) contains fewer than five digits. If a record does contain a ZIP Code of fewer than five digits, the record is excluded from the mail merge and the address is marked as invalid.


```vb
Sub CheckRecords() 
 
 Dim intCount As Integer 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included record 
 ' in the data source 
 .ActiveRecord = wdFirstRecord 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that field six must be greater than 
 'or equal to five 
 If Len(.DataFields(6).Value) < 5 Then 
 
 'Exclude the record if field six is less than five 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record 
 'explaining why the record was excluded 
 'from the mail merge 
 .InvalidComments = "The ZIP Code for this record " &; _ 
 "has fewer than five digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = wdNextRecord 
 
 'End the loop when the counter variable equals the 
 'number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


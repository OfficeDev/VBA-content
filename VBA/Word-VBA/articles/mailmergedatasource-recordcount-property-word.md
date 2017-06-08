---
title: MailMergeDataSource.RecordCount Property (Word)
keywords: vbawd10.chm152895500
f1_keywords:
- vbawd10.chm152895500
ms.prod: word
api_name:
- Word.MailMergeDataSource.RecordCount
ms.assetid: d69db5d2-7ef0-dd9a-7e03-0029f6defd37
ms.date: 06/08/2017
---


# MailMergeDataSource.RecordCount Property (Word)

Returns a  **Long** that represents the number of records in the data source. Read-only.


## Syntax

 _expression_ . **RecordCount**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

If Microsoft Word cannot determine the number of records in a data source, the  **RecordCount** property will return a value of -1.


## Example

This example loops through the records in the data source and verifies that the postal code field (field six in this example) is not fewer than five digits. If it is, it removes the record from the mail merge. If you want to make sure that the locator code is added to the postal code, you can change the length value from 5 to 10. Therefore, if a postal code is fewer than ten digits it will be removed from the mail merge.


```vb
Sub ExcludeRecords() 
 
 On Error GoTo ErrorHandler 
 
 With ActiveDocument.MailMerge.DataSource 
 .ActiveRecord = wdFirstRecord 
 Do 
 
 'Counts the number of digits in the postal code field and if 
 'it is fewer than 5, the record is excluded from the mail merge, 
 'marked as having an invalid address, and given a comment 
 'describing why the postal code was removed 
 If Len(.DataFields(6).Value) < 5 Then 
 .Included = False 
 .InvalidAddress = True 
 .InvalidComments = "The ZIP Code for this record" &; _ 
 "has fewer than five digits. This record will be" &; _ 
 "removed from the mail merge process." 
 End If 
 If .ActiveRecord <> .RecordCount Then 
 .ActiveRecord = wdNextRecord 
 End If 
 Loop Until .ActiveRecord = .RecordCount 
ErrorHandler: 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


---
title: MailMergeDataSource.ActiveRecord Property (Publisher)
keywords: vbapb10.chm6291459
f1_keywords:
- vbapb10.chm6291459
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.ActiveRecord
ms.assetid: 0f092eb4-6e65-9235-83e2-a04b813b2390
ms.date: 06/08/2017
---


# MailMergeDataSource.ActiveRecord Property (Publisher)

Returns or sets a  **Long** that represents the active mail merge record. Read/write.


## Syntax

 _expression_. **ActiveRecord**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Long


## Remarks

The active record number is the position of the record in the query result produced by the current query options; as such, this number is not necessarily the position of the record in the data source.


## Example

This example validates that the value entered into the PostalCode field is ten characters long (U.S. ZIP Code plus 4-digit locator code). If it is not, it is excluded from the mail merge and marked with a comment.


```vb
Sub ValidateZip() 
 
 Dim intCount As Integer 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included 
 'record in the data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that the PostalCode field 
 'must be greater than or equal to ten digits 
 If Len(.DataFields.Item("PostalCode").Value) < 10 Then 
 
 'Exclude the record if the PostalCode field 
 'is less than ten digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP code for this record is " _ 
 &; "less than ten digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 
 
End Sub
```



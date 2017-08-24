---
title: MailMergeDataSource.RecordCount Property (Publisher)
keywords: vbapb10.chm6291477
f1_keywords:
- vbapb10.chm6291477
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.RecordCount
ms.assetid: 56b929bf-9b7f-dd83-98b7-35bf96028732
ms.date: 06/08/2017
---


# MailMergeDataSource.RecordCount Property (Publisher)

Returns a  **Long** that represents the number of records in the data source. Read-only.


## Syntax

 _expression_. **RecordCount**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Long


## Example

This example validates ZIP Codes in the attached data source for five digits. If the length of the ZIP Code is fewer than five digits, the record is excluded from the mail merge process. This example assumes the postal codes are U.S. ZIP Codes. You could modify this example to search for ZIP Codes that have a 4-digit locator code appended to the ZIP Code, and then exclude all records that do not contain the locator code.


```vb
Sub Validate 
 
 Dim intCount As Integer 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included record in the 
 'data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that field six must be greater than or 
 'equal to five digits in length 
 If Len(.DataFields.Item(6).Value) < 5 Then 
 
 'Exclude the record if field six contains fewer than five digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP Code for this record has " _ 
 &; "fewer than five digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 

```



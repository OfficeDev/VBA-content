---
title: MailMergeDataSource.FindRecord Method (Publisher)
keywords: vbapb10.chm6291480
f1_keywords:
- vbapb10.chm6291480
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.FindRecord
ms.assetid: a4b37255-bdff-ac61-6d18-05a4fe008beb
ms.date: 06/08/2017
---


# MailMergeDataSource.FindRecord Method (Publisher)

Searches the contents of the specified mail merge data source for text in a particular field. Returns a  **Boolean** indicating whether the search text is found; **True** if the search text is found.


## Syntax

 _expression_. **FindRecord**( **_FindText_**,  **_Field_**)

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FindText|Required| **String**|The text to look for.|
|Field|Optional| **String**|The name of the field to be searched.|

### Return Value

Boolean


## Example

This example displays a merge publication for the first record in which the FirstName field contains Joe. If the record is found, the record number is stored in a variable.


```vb
Sub FindDataSourceRecord() 
 Dim dsMain As MailMergeDataSource 
 Dim intRecord As Integer 
 
 'Makes the data in the data source records instead of the field codes 
 ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
 
 Set dsMain = ActiveDocument.MailMerge.DataSource 
 
 If dsMain.FindRecord(FindText:="Joe", _ 
 Field:="FirstName") = True Then 
 intRecord = dsMain.ActiveRecord 
 End If 
 
End Sub
```



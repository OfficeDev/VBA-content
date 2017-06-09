---
title: MailMergeDataSource.FindRecord Method (Word)
keywords: vbawd10.chm152895590
f1_keywords:
- vbawd10.chm152895590
ms.prod: word
api_name:
- Word.MailMergeDataSource.FindRecord
ms.assetid: 1d4bc94c-8305-57d9-d63f-ce4ac54aa4d4
ms.date: 06/08/2017
---


# MailMergeDataSource.FindRecord Method (Word)

Searches the contents of the specified mail merge data source for text in a particular field. Returns  **True** if the search text is found. **Boolean** .


## Syntax

 _expression_ . **FindRecord**( **_FindText_** , **_Field_** )

 _expression_ Required. A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindText_|Required| **String**|The text to be looked for.|
| _Field_|Required| **Variant**|The name of the field to be searched.|

### Return Value

Boolean


## Remarks

This method corresponds to the  **Find Record** button on the **Mail Merge** toolbar.

The  **FindRecord** method does a forward search only. Therefore, if the active record is not the first record in the data source and the record for which you are searching is before the active record, the **FindRecord** method will return no results. To ensure that the entire data source is searched, set the **ActiveRecord** property to **wdFirstRecord** .


## Example

This example displays a merge document for the first record in which the FirstName field contains "Joe." If the record is found, the number of the record is stored in the numRecord variable.


```vb
Dim dsMain As MailMergeDataSource 
Dim numRecord As Integer 
 
ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
Set dsMain = ActiveDocument.MailMerge.DataSource 
If dsMain.FindRecord(FindText:="Joe", _ 
 Field:="FirstName") = True Then 
 numRecord = dsMain.ActiveRecord 
End If
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


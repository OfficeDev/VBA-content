---
title: MailMergeDataSource.SetAllErrorFlags Method (Word)
keywords: vbawd10.chm152895592
f1_keywords:
- vbawd10.chm152895592
ms.prod: word
api_name:
- Word.MailMergeDataSource.SetAllErrorFlags
ms.assetid: 9419781e-ca05-dac7-d11f-91e002a6cb84
ms.date: 06/08/2017
---


# MailMergeDataSource.SetAllErrorFlags Method (Word)

Marks all records in a mail merge data source as containing invalid data in an address field.


## Syntax

 _expression_ . **SetAllErrorFlags**( **_Invalid_** , **_InvalidComment_** )

 _expression_ Required. A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Invalid_|Required| **Boolean**| **True** marks all records in the data source of a mail merge as invalid.|
| _InvalidComment_|Required| **String**|Text describing the invalid setting.|

## Remarks

You can individually mark data source records that contain invalid data in an address field by using the  **InvalidAddress** and **InvalidComments** properties.


## Example

This example marks all records in the data source as containing an invalid address field, sets a comment as to why it is invalid, and excludes all records from the mail merge.


```vb
Sub FlagAllRecords() 
 With ActiveDocument.MailMerge.DataSource 
 .SetAllErrorFlags Invalid:=True, InvalidComment:= _ 
 "All records in the data source have only 5-" _ 
 &; "digit ZIP Codes. Need 5+4 digit ZIP Codes." 
 .SetAllIncludedFlags Included:=False 
 End With 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)


---
title: MailMergeDataSource.SetAllIncludedFlags Method (Word)
keywords: vbawd10.chm152895591
f1_keywords:
- vbawd10.chm152895591
ms.prod: word
api_name:
- Word.MailMergeDataSource.SetAllIncludedFlags
ms.assetid: 1fd70215-9b74-bf36-7ba2-9c02e2dc6a89
ms.date: 06/08/2017
---


# MailMergeDataSource.SetAllIncludedFlags Method (Word)

Includes or excludes flagged records in a data source from a mail merge.


## Syntax

 _expression_ . **SetAllIncludedFlags**( **_Included_** )

 _expression_ Required. A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Included_|Required| **Boolean**| **True** to include all data source records in a mail merge. **False** to exclude all data source records from a mail merge.|

## Remarks

You can set individual records in a data source to be included in or excluded from a mail merge using the  **Included** property.


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


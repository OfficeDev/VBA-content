---
title: MailMergeDataSource.SetSortOrder Method (Publisher)
keywords: vbapb10.chm6291489
f1_keywords:
- vbapb10.chm6291489
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.SetSortOrder
ms.assetid: 0ecb5f77-2cd1-92c6-b7f2-bf709f015ba5
ms.date: 06/08/2017
---


# MailMergeDataSource.SetSortOrder Method (Publisher)

Sets the sort order for mail merge data.


## Syntax

 _expression_. **SetSortOrder**( **_SortField1_**,  **_SortAscending1_**,  **_SortField2_**,  **_SortAscending2_**,  **_SortField3_**,  **_SortAscending3_**)

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SortField1|Optional| **String**|The first field on which to sort the mail merge data. Default is an empty string.|
|SortAscending1|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField1; **False** to perform a descending sort.|
|SortField2|Optional| **String**|The second field on which to sort the mail merge data. Default is an empty string.|
|SortAscending2|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField2; **False** to perform a descending sort.|
|SortField3|Optional| **String**|The third field on which to sort the mail merge data. Default is an empty string.|
|SortAscending3|Optional| **Boolean**| **True** (default) to perform an ascending sort on SortField3; **False** to perform a descending sort.|

## Example

The following example sorts mail merge data first on postal code in descending order, then on last name and first name in ascending order.


```vb
ActiveDocument.MailMerge.DataSource.SetSortOrder _ 
 SortField1:="ZIPCode", SortAscending1:=False, _ 
 SortField2:="LastName", SortField3:="FirstName"
```



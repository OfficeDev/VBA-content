---
title: SparklineGroups.Group Method (Excel)
keywords: vbaxl10.chm869080
f1_keywords:
- vbaxl10.chm869080
ms.prod: excel
api_name:
- Excel.SparklineGroups.Group
ms.assetid: a5e01669-1922-4b26-158d-3c3aa70a101a
ms.date: 06/08/2017
---


# SparklineGroups.Group Method (Excel)

Groups the selected sparklines.


## Syntax

 _expression_ . **Group**( **_Location_** )

 _expression_ A variable that represents a **[SparklineGroups](sparklinegroups-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **[Range](range-object-excel.md)**|The location of the first cell in the group.|

### Return Value

Nothing


## Example

This example selects the range A1:A4 and groups the sparklines in that range.


```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Group Location:=Range("A1")
```


## See also


#### Concepts


[SparklineGroups Object](sparklinegroups-object-excel.md)


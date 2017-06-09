---
title: SparklineGroup.ModifySourceData Method (Excel)
keywords: vbaxl10.chm871080
f1_keywords:
- vbaxl10.chm871080
ms.prod: excel
api_name:
- Excel.SparklineGroup.ModifySourceData
ms.assetid: 35c1c1ed-b61d-2412-961f-8eb74b5563a2
ms.date: 06/08/2017
---


# SparklineGroup.ModifySourceData Method (Excel)

Sets the range that represents the source data for the sparkline group.


## Syntax

 _expression_ . **ModifySourceData**( **_SourceData_** )

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceData_|Required| **String**|The range that represents the source data.|

### Return Value

Nothing


## Example

This example selects a sparkline group in the location A1:A4 and modifies the source data to include an additional column using the data in the range B1:D4.


```vb
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifySourceData "B1:D4"
```


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)


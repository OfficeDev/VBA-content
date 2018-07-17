---
title: SparklineGroup.ModifyDateRange Method (Excel)
keywords: vbaxl10.chm871082
f1_keywords:
- vbaxl10.chm871082
ms.prod: excel
api_name:
- Excel.SparklineGroup.ModifyDateRange
ms.assetid: 2de21c82-64b6-6095-0c47-cd20354d9739
ms.date: 06/08/2017
---


# SparklineGroup.ModifyDateRange Method (Excel)

Sets the date range for the sparkline group.


## Syntax

 _expression_ . **ModifyDateRange**( **_DateRange_** )

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DateRange_|Required| **String**|The date range for the sparkline group.|

### Return Value

Nothing


## Example

This example selects a sparkline group in the location A2:A5 and sets the date range equal to B1:E1. If the cells in range B1:E1 do not contain date values the data is not displayed.


```vb
Range("A2:A5").Select 
ActiveCell.SparklineGroups.Item(1).ModifyDateRange "Sheet1!B1:E1"
```


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)


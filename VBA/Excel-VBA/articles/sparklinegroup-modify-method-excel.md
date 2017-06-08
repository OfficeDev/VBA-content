---
title: SparklineGroup.Modify Method (Excel)
keywords: vbaxl10.chm871081
f1_keywords:
- vbaxl10.chm871081
ms.prod: excel
api_name:
- Excel.SparklineGroup.Modify
ms.assetid: 596cdecb-dd03-0a63-e2b8-9aa459ff719c
ms.date: 06/08/2017
---


# SparklineGroup.Modify Method (Excel)

Sets the location and the source data for the sparkline group.


## Syntax

 _expression_ . **Modify**( **_Location_** , **_SourceData_** )

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **Range**|The  **[Range](range-object-excel.md)** object that represents the location of the sparkline group.|
| _SourceData_|Required| **String**|The range that represents the source data for the sparkline group.|

### Return Value

Nothing


## Example

This examples selects a sparkline group in the location A1:A4 and removes a row of data by changing the sparkline group location to equal A1:A3. The data source must also be modified to only include the first three rows of data.


```vb
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).Modify Location:=Range("$A$1:$A$3"), SourceData:="Sheet1!B1:D3"
```


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)


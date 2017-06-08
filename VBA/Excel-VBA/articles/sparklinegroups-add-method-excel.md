---
title: SparklineGroups.Add Method (Excel)
keywords: vbaxl10.chm869073
f1_keywords:
- vbaxl10.chm869073
ms.prod: excel
api_name:
- Excel.SparklineGroups.Add
ms.assetid: ae41a572-c073-5251-b2c1-884e832e8ae5
ms.date: 06/08/2017
---


# SparklineGroups.Add Method (Excel)

Creates a new sparkline group and returns a  **[SparklineGroup](sparklinegroup-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_Type_** , **_SourceData_** )

 _expression_ A variable that represents a **SparklineGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **XlSparkType**|The type of sparkline.|
| _SourceData_|Required| **String**|Represents the range to use to create the sparkline.|

### Return Value

SparklineGroup


## Example

This example adds a sparkline group to the range A1:A4. The sparklines in the group are column sparklines and are bound to the data in the range B1:E4.


```vb
Range("$A$1:$A$4").SparklineGroups.Add Type:=xlSparkColumn, SourceData:= _ 
 "Sheet2!B1:E4"
```


## See also


#### Concepts


[SparklineGroups Object](sparklinegroups-object-excel.md)


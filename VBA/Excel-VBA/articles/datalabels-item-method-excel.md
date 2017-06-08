---
title: DataLabels.Item Method (Excel)
keywords: vbaxl10.chm584106
f1_keywords:
- vbaxl10.chm584106
ms.prod: excel
api_name:
- Excel.DataLabels.Item
ms.assetid: bc45ebcc-00f0-c253-0d68-002d8f20d750
ms.date: 06/08/2017
---


# DataLabels.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **DataLabels** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

### Return Value

A  **[DataLabel](datalabel-object-excel.md)** object contained by the collection.


## Example

This example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.


```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000"
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)


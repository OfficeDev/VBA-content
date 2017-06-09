---
title: Series.Points Method (Excel)
keywords: vbaxl10.chm578104
f1_keywords:
- vbaxl10.chm578104
ms.prod: excel
api_name:
- Excel.Series.Points
ms.assetid: 9b6f08a1-3fbe-e9bc-a509-345a3d2d78b3
ms.date: 06/08/2017
---


# Series.Points Method (Excel)

Returns an object that represents a single point (a  **[Point](point-object-excel.md)** object) or a collection of all the points (a **[Points](points-object-excel.md)** collection) in the series. Read-only


## Syntax

 _expression_ . **Points**( **_Index_** )

 _expression_ An expression that returns a **Series** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the point.|

### Return Value

Object


## Example

This example applies a data label to point one in series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1).Points(1).ApplyDataLabels
```


## See also


#### Concepts


[Series Object](series-object-excel.md)


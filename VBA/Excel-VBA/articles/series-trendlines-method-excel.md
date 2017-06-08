---
title: Series.Trendlines Method (Excel)
keywords: vbaxl10.chm578107
f1_keywords:
- vbaxl10.chm578107
ms.prod: excel
api_name:
- Excel.Series.Trendlines
ms.assetid: d42609e1-011c-6cb3-286d-192284cd8ab8
ms.date: 06/08/2017
---


# Series.Trendlines Method (Excel)

Returns an object that represents a single trendline (a  **[Trendline](trendline-object-excel.md)** object) or a collection of all the trendlines (a **[Trendlines](trendlines-object-excel.md)** collection) for the series.


## Syntax

 _expression_ . **Trendlines**( **_Index_** )

 _expression_ A variable that represents a **Series** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the trendline.|

### Return Value

Object


## Example

This example adds a linear trendline to series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1).Trendlines.Add Type:=xlLinear
```


## See also


#### Concepts


[Series Object](series-object-excel.md)


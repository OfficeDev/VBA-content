---
title: ChartGroup.SeriesCollection Method (Excel)
keywords: vbaxl10.chm568088
f1_keywords:
- vbaxl10.chm568088
ms.prod: excel
api_name:
- Excel.ChartGroup.SeriesCollection
ms.assetid: 7da987dc-5629-1b7d-9269-cadbec2f8c46
ms.date: 06/08/2017
---


# ChartGroup.SeriesCollection Method (Excel)

Returns an object that represents either a single series (a  **[Series](series-object-excel.md)** object) or a collection of all the series (a **[SeriesCollection](seriescollection-object-excel.md)** collection) in the chart or chart group.


## Syntax

 _expression_ . **SeriesCollection**( **_Index_** )

 _expression_ A variable that represents a **ChartGroup** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the series.|

### Return Value

Object


## Example

This example turns on data labels for series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1).HasDataLabels = True
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


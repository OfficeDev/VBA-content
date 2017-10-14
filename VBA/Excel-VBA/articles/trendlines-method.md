---
title: Trendlines Method
keywords: vbagr10.chm3077635
f1_keywords:
- vbagr10.chm3077635
ms.prod: excel
api_name:
- Excel.Trendlines
ms.assetid: 2379333d-1cca-bd04-2dec-170bd5d40f67
ms.date: 06/08/2017
---


# Trendlines Method

Returns an object that represents a single trendline or a collection of all the trendlines for the series.

 _expression_. **Trendlines**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The name or number of the trendline.

## Example

This example adds a linear trendline to series one.


```
myChart.SeriesCollection(1).Trendlines.Add Type:=xlLinear
```



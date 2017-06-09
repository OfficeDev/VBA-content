---
title: Trendline Object
keywords: vbagr10.chm131198
f1_keywords:
- vbagr10.chm131198
ms.prod: excel
api_name:
- Excel.Trendline
ms.assetid: 227bc97a-1bdf-f90b-9bef-f9f611c643af
ms.date: 06/08/2017
---


# Trendline Object

Represents a trendline in the specified chart. A trendline shows the trend, or direction, of data in a series. The  **Trendline** object is a member of the **[Trendlines](trendlines-collection-excel.md)** collection, which contains all the  **Trendline** objects for a single series.


## Using the Trendline Object

Use  **Trendlines**( _index_), where  _index_ is the trendline's index number, to return a single **Trendline** object. The following example changes the trendline type for series one in the chart. If the series has no trendline, this example will fail.


```
myChart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```

The index number denotes the order in which the trendlines are added to the series.  `Trendlines(1)` is the first trendline added to the series, and is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.



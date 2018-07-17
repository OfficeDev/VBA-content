---
title: Trendlines Collection (Excel)
keywords: vbagr10.chm5208077
f1_keywords:
- vbagr10.chm5208077
ms.prod: excel
ms.assetid: 4b12461a-65a2-c535-e98d-ff68ffa5919c
ms.date: 06/08/2017
---


# Trendlines Collection (Excel)

A collection of all the  **[Trendline](trendline-object.md)** objects for the specified series. Each  **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Using the Trendlines Collection

Use the  **Trendlines** method to return the **Trendlines** collection. The following example displays the number of trendlines for series one in the chart.


```vb
MsgBox myChart.SeriesCollection(1).Trendlines.Count
```

Use the  **[Add](add-method.md)** method to create a new trendline and add it to the series. The following example adds a linear trendline to series one in the chart.




```vb
With myChart.SeriesCollection(1).Trendlines 
 .Add Type:=xlLinear, Name:="Linear Trend" 
End With
```

Use  **Trendlines**( _index_), where  _index_ is the trendline's index number, to return a single **TrendLine** object. The following example changes the trendline type for series one in the chart. If the series has no trendline, this example will fail.




```
myChart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```

The index number denotes the order in which the trendlines are added to the series.  `Trendlines(1)` is the first trendline added to the series, and is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.



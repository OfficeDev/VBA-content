---
title: Series Object
keywords: vbagr10.chm131115
f1_keywords:
- vbagr10.chm131115
ms.prod: excel
api_name:
- Excel.Series
ms.assetid: c4446d04-9a3a-4f95-7b3f-adaf1ad2252c
ms.date: 06/08/2017
---


# Series Object

Represents a series in the specified chart. The  **Series** object is a member of the **[SeriesCollection](seriescollection-collection-excel.md)** collection.


## Using the Series Object

Use  **SeriesCollection**( _index_), where  _index_ is the series' index number or name, to return a single **Series** object. The following example sets the color of the interior for series one in the chart.


```
myChart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```

The series index number indicates the order in which the series are added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.



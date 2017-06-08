---
title: Point Object
keywords: vbagr10.chm131114
f1_keywords:
- vbagr10.chm131114
ms.prod: excel
api_name:
- Excel.Point
ms.assetid: 944d5edb-b1e7-7aed-5ead-bde3878b26e5
ms.date: 06/08/2017
---


# Point Object

Represents a single point in a series on the specified chart. The  **Point** object is a member of the **[Points](points-collection-excel.md)** collection, which contains all the points in the specified series.


## Using the Point Object

Use  **Points**( _index_), where  _index_ is the point's index number, to return a single **Point** object. Points are numbered from left to right in the series. `Points(1)` is the leftmost point, and is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one. For this example to work, series one must be a 2-D line, scatter, or radar series.


```
myChart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```



---
title: Points Collection (Excel)
keywords: vbagr10.chm5207812
f1_keywords:
- vbagr10.chm5207812
ms.prod: excel
ms.assetid: b41c8f08-880e-1f4a-0456-3f77c0741bc6
ms.date: 06/08/2017
---


# Points Collection (Excel)

A collection of all the  **[Point](point-object.md)** objects in the specified series in a chart.


## Using the Points Collection

Use the  **Points** method to return the **Points** collection. The following example adds a data label to the last point in series one in the chart.


```vb
Dim pts As Points 
Set pts = myChart.SeriesCollection(1).Points 
pts(pts.Count).ApplyDataLabels Type:=xlShowValue
```

Use  **Points**( _index_), where  _index_ is the point's index number, to return a single **Point** object. Points are numbered from left to right in the series. `Points(1)` is the leftmost point, and is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one in the chart. The specified series must be a 2-D line, scatter, or radar series.




```
myChart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```



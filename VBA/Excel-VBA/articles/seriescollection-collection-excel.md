---
title: SeriesCollection Collection (Excel)
keywords: vbagr10.chm131116
f1_keywords:
- vbagr10.chm131116
ms.prod: excel
ms.assetid: c5d00466-f7a1-7e6f-56e4-958901dbe3e3
ms.date: 06/08/2017
---


# SeriesCollection Collection (Excel)

A collection of all the  **[Series](series-object.md)** objects in the specified chart or chart group.


## Using the SeriesCollection Collection

Use the  **SeriesCollection** method to return the **SeriesCollection** collection. The following example adjusts the interior color for each series in the collection:


```vb
For X = 1 To myChart.SeriesCollection.Count 
 With myChart.SeriesCollection(X) 
 .Interior.Color = RGB(X * 75, 50, X * 50) 
 End With 
Next X
```

Use  **SeriesCollection**( _index_), where  _index_ is the series' index number or name, to return a single **Series** object. The following example sets the color of the interior for series one in the chart to red.




```
myChart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```



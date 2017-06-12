---
title: Trendline Object (Excel)
keywords: vbaxl10.chm593072
f1_keywords:
- vbaxl10.chm593072
ms.prod: excel
api_name:
- Excel.Trendline
ms.assetid: 5c04b065-57f4-a059-7c22-50612bd727ea
ms.date: 06/08/2017
---


# Trendline Object (Excel)

Represents a trendline in a chart.


## Remarks

A trendline shows the trend, or direction, of data in a series. The  **Trendline** object is a member of the **[Trendlines](trendlines-object-excel.md)** collection. The **Trendlines** collection contains all the **Trendline** objects for a single series.


## Example

Use  **Trendlines** ( _index_), where  _index_ is the trendline index number, to return a single **Trendline** object.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.

The following example changes the trendline type for the first series in embedded chart one on worksheet one. If the series has no trendline, this example will fail.




```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)



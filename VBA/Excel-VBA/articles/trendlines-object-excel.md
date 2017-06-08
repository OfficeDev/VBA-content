---
title: Trendlines Object (Excel)
keywords: vbaxl10.chm591072
f1_keywords:
- vbaxl10.chm591072
ms.prod: excel
api_name:
- Excel.Trendlines
ms.assetid: 752cde45-c628-7550-6c88-07405821e348
ms.date: 06/08/2017
---


# Trendlines Object (Excel)

A collection of all the  **[Trendline](trendline-object-excel.md)** objects for the specified series.


## Remarks

Each  **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Example

Use the  **[Trendlines](series-trendlines-method-excel.md)** method to return the **Trendlines** collection. The following example displays the number of trendlines for series one in Chart1.


```
MsgBox Charts(1).SeriesCollection(1).Trendlines.Count
```

Use the  **[Add](trendlines-add-method-excel.md)** method to create a new trendline and add it to the series. The following example adds a linear trendline to the first series in embedded chart one on Sheet1.




```
Worksheets("sheet1").ChartObjects(1).Chart.SeriesCollection(1) _ 
 .Trendlines.Add type:=xlLinear, name:="Linear Trend"
```

Use  **Trendlines** ( _index_), where  _index_ is the trendline index number, to return a single **TrendLine** object. The following example changes the trendline type for the first series in embedded chart one on worksheet one. If the series has no trendline, this example will fail.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.




```
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```


## Methods



|**Name**|
|:-----|
|[Add](trendlines-add-method-excel.md)|
|[Item](trendlines-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](trendlines-application-property-excel.md)|
|[Count](trendlines-count-property-excel.md)|
|[Creator](trendlines-creator-property-excel.md)|
|[Parent](trendlines-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

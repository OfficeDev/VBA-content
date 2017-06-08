---
title: Points Object (Excel)
keywords: vbaxl10.chm573072
f1_keywords:
- vbaxl10.chm573072
ms.prod: excel
api_name:
- Excel.Points
ms.assetid: 918dc385-ed61-262e-033f-ba829f5ee8b2
ms.date: 06/08/2017
---


# Points Object (Excel)

A collection of all the  **[Point](point-object-excel.md)** objects in the specified series in a chart.


## Remarks

Use  **[Points](series-points-method-excel.md)** ( _index_ ), where _index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point.


## Example

Use the  **Points** method to return the **[Points](points-object-excel.md)** collection. The following example adds a data label to the last point on series one in embedded chart one on worksheet one.


```
Dim pts As Points 
Set pts = Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points 
pts(pts.Count).ApplyDataLabels type:=xlShowValue
```

 The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2-D line, scatter, or radar series.




```
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```


## Methods



|**Name**|
|:-----|
|[Item](points-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](points-application-property-excel.md)|
|[Count](points-count-property-excel.md)|
|[Creator](points-creator-property-excel.md)|
|[Parent](points-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

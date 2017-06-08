---
title: ChartGroups Method
keywords: vbagr10.chm65544
f1_keywords:
- vbagr10.chm65544
ms.prod: excel
api_name:
- Excel.ChartGroups
ms.assetid: e25258c1-14d4-bb0c-b442-f6c811b19847
ms.date: 06/08/2017
---


# ChartGroups Method

Returns an object that represents either a single chart group or a collection of all the chart groups in the chart. The returned collection includes every type of group.

 _expression_. **ChartGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The chart group number.

## Example

This example turns on up and down bars for chart group one and then sets their colors. The example should be run on a 2-D line chart containing two series that intersect at one or more data points.


```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



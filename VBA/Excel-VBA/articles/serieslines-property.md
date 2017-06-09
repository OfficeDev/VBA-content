---
title: SeriesLines Property
keywords: vbagr10.chm5207974
f1_keywords:
- vbagr10.chm5207974
ms.prod: excel
api_name:
- Excel.SeriesLines
ms.assetid: ebfea917-8678-7d05-df9d-2102f396ea59
ms.date: 06/08/2017
---


# SeriesLines Property

Returns a  **[SeriesLines](serieslines-object.md)** object that represents the series lines for the specified stacked bar chart or stacked column chart. Applies only to stacked bar and stacked column charts. Read-only.


## Example

This example turns on series lines for chart group one on the chart and then sets their line style, weight, and color. The example should be run on a 2-D stacked column chart that has two or more series.


```vb
With myChart.ChartGroups(1) 
 .HasSeriesLines = True 
 With .SeriesLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```



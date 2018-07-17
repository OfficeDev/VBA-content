---
title: HasSeriesLines Property
keywords: vbagr10.chm65601
f1_keywords:
- vbagr10.chm65601
ms.prod: excel
api_name:
- Excel.HasSeriesLines
ms.assetid: fd101b78-4499-31bd-1243-47738c1eb00f
ms.date: 06/08/2017
---


# HasSeriesLines Property

 **True** if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to stacked column charts, bar charts, Pie of Pie charts, or Bar of Pie charts. Read/write **Boolean**.


## Example

This example turns on series lines for chart group one and then sets their line style, weight, and color. The example should be run on a 2-D stacked column chart that has two or more series.


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



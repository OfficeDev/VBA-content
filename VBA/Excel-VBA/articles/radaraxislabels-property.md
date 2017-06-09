---
title: RadarAxisLabels Property
keywords: vbagr10.chm65680
f1_keywords:
- vbagr10.chm65680
ms.prod: excel
api_name:
- Excel.RadarAxisLabels
ms.assetid: e382e92c-96f2-a9ee-720f-dcb85e5e2e7c
ms.date: 06/08/2017
---


# RadarAxisLabels Property

Returns a  **[TickLabels](ticklabels-object.md)** object that represents the radar axis labels for the specified chart group. Read-only.


## Example

This example turns on radar axis labels for chart group one on the chart and then sets the color for the labels. The example should be run on a radar chart.


```vb
With myChart.ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```



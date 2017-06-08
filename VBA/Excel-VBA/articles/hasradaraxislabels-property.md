---
title: HasRadarAxisLabels Property
keywords: vbagr10.chm65600
f1_keywords:
- vbagr10.chm65600
ms.prod: excel
api_name:
- Excel.HasRadarAxisLabels
ms.assetid: 8baa636a-262c-15b4-f8d5-94d77a8101c5
ms.date: 06/08/2017
---


# HasRadarAxisLabels Property

 **True** if a radar chart has axis labels. Applies only to radar charts. Read/write **Boolean**.


## Example

This example turns on radar axis labels for chart group one and sets their color. The example should be run on a radar chart.


```vb
With myChart.ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```



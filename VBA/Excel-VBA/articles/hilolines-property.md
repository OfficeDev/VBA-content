---
title: HiLoLines Property
keywords: vbagr10.chm65679
f1_keywords:
- vbagr10.chm65679
ms.prod: excel
api_name:
- Excel.HiLoLines
ms.assetid: ed2ff722-b477-4346-d807-3d2615abd845
ms.date: 06/08/2017
---


# HiLoLines Property

Returns a  **[HiLoLines](hilolines-object.md)** object that represents the high-low lines for the specified series on a line chart. Applies only to line charts. Read-only.


## Example

This example turns on high-low lines for chart group one on the chart and then sets their line style, weight, and color. The example should be run on a 2-D line chart that has three series of stock-quote-like data (high-low-close).


```vb
With myChart.ChartGroups(1) 
 .HasHiLoLines = True 
 With .HiLoLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```



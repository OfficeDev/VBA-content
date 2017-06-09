---
title: DropLines Property
keywords: vbagr10.chm5207331
f1_keywords:
- vbagr10.chm5207331
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 13dd4b80-669e-94c1-d592-439129d42d56
ms.date: 06/08/2017
---


# DropLines Property

Returns a  **[DropLines](droplines-object.md)** object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.


## Example

This example turns on drop lines for chart group one and then sets their line style, weight, and color. The example should be run on a 2-D line chart that has one series.


```vb
With myChart.ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```



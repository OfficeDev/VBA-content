---
title: DownBars Property
keywords: vbagr10.chm5207325
f1_keywords:
- vbagr10.chm5207325
ms.prod: excel
api_name:
- Excel.DownBars
ms.assetid: 752b1b94-9027-876a-54a2-7aabed4e055b
ms.date: 06/08/2017
---


# DownBars Property

Returns a  **[DownBars](downbars-object.md)** object that represents the down bars on a line chart. Applies only to line charts. Read-only.


## Example

This example turns on up bars and down bars for chart group one and then sets their colors. The example should be run on a 2-D line chart that has two series that cross each other at one or more data points.


```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



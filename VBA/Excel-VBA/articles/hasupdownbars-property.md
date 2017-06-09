---
title: HasUpDownBars Property
keywords: vbagr10.chm5207518
f1_keywords:
- vbagr10.chm5207518
ms.prod: excel
api_name:
- Excel.HasUpDownBars
ms.assetid: c3785986-a013-727c-95e6-56a732b8b40f
ms.date: 06/08/2017
---


# HasUpDownBars Property

 **True** if the specified line chart has up and down bars. Applies only to line charts. Read/write **Boolean**.


## Example

This example turns on up and down bars for chart group one and then sets their colors. The example should be run on a 2-D line chart containing two series that cross each other at one or more data points.


```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```



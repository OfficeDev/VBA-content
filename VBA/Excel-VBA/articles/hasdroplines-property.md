---
title: HasDropLines Property
keywords: vbagr10.chm5207474
f1_keywords:
- vbagr10.chm5207474
ms.prod: excel
api_name:
- Excel.HasDropLines
ms.assetid: 31f00864-86bc-9237-bf93-b52ab8cd1b59
ms.date: 06/08/2017
---


# HasDropLines Property

 **True** if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write **Boolean**.


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



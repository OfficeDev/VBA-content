---
title: HasHiLoLines Property
keywords: vbagr10.chm5207483
f1_keywords:
- vbagr10.chm5207483
ms.prod: excel
api_name:
- Excel.HasHiLoLines
ms.assetid: 57018e82-acf1-039f-3fa5-d2319385c3d5
ms.date: 06/08/2017
---


# HasHiLoLines Property

 **True** if the line chart has high-low lines. Applies only to line charts. Read/write **Boolean**.


## Example

This example turns on high-low lines for chart group one and then sets line style, weight, and color. The example should be run on a 2-D line chart that has three series of stock-quote-like data (high-low-close).


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



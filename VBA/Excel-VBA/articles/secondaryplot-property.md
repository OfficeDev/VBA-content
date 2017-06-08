---
title: SecondaryPlot Property
keywords: vbagr10.chm5207958
f1_keywords:
- vbagr10.chm5207958
ms.prod: excel
api_name:
- Excel.SecondaryPlot
ms.assetid: 6806a9d3-06cc-3786-5d1e-fbc23680da7a
ms.date: 06/08/2017
---


# SecondaryPlot Property

 **True** if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts. Read/write **Boolean**.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example moves point four to the secondary section of the chart.


```vb
With myChart.SeriesCollection(1) 
 .Points(4).SecondaryPlot = True 
End With
```



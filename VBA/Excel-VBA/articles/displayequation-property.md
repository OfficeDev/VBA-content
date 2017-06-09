---
title: DisplayEquation Property
keywords: vbagr10.chm5207312
f1_keywords:
- vbagr10.chm5207312
ms.prod: excel
api_name:
- Excel.DisplayEquation
ms.assetid: f3638bfd-d25d-96b4-5c20-2acf8703658d
ms.date: 06/08/2017
---


# DisplayEquation Property

 **True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to **True** automatically turns on data labels. Read/write **Boolean**.


## Example

This example displays the R-squared value and equation for trendline one. The example should be run on a 2-D column chart that has a trendline for the first series.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
End With
```



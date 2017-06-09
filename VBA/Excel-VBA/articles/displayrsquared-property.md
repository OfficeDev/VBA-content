---
title: DisplayRSquared Property
keywords: vbagr10.chm5207314
f1_keywords:
- vbagr10.chm5207314
ms.prod: excel
api_name:
- Excel.DisplayRSquared
ms.assetid: cc8ac282-19b1-00d8-14a7-738f5574f1cb
ms.date: 06/08/2017
---


# DisplayRSquared Property

 **True** if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to **True** automatically turns on data labels. Read/write **Boolean**.


## Example

This example displays the R-squared value and equation for trendline one. The example should be run on a 2-D column chart that has a trendline for the first series.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
End With
```



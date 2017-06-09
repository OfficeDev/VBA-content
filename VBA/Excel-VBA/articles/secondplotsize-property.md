---
title: SecondPlotSize Property
keywords: vbagr10.chm5207961
f1_keywords:
- vbagr10.chm5207961
ms.prod: excel
api_name:
- Excel.SecondPlotSize
ms.assetid: 43d450e1-0ef0-fd51-fbf1-b07742217fc9
ms.date: 06/08/2017
---


# SecondPlotSize Property

Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 through 200. Read/write  **Long**.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. The secondary section is 50 percent of the size of the primary pie.


```vb
With myChart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
 .SecondPlotSize = 50 
End With
```



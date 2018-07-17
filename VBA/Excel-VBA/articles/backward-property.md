---
title: Backward Property
keywords: vbagr10.chm65721
f1_keywords:
- vbagr10.chm65721
ms.prod: excel
api_name:
- Excel.Backward
ms.assetid: a92f33cb-45cd-baea-57e1-d76f44b041cb
ms.date: 06/08/2017
---


# Backward Property

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write  **Long**.


## Example

This example sets the number of units that the trendline extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```



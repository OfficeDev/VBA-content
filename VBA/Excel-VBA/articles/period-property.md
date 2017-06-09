---
title: Period Property
keywords: vbagr10.chm65720
f1_keywords:
- vbagr10.chm65720
ms.prod: excel
api_name:
- Excel.Period
ms.assetid: 6f0378a3-a158-b21d-eef3-acde9e86f94b
ms.date: 06/08/2017
---


# Period Property

Returns or sets the period for the moving-average trendline. Read/write  **Long**.


## Example

This example sets the period for the moving-average trendline. The example should be run on a 2-D column chart with a single series that contains 10 data points and a moving-average trendline.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 If .Type = xlMovingAvg Then .Period = 5 
End With
```



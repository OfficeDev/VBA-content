---
title: Forward Property
keywords: vbagr10.chm65727
f1_keywords:
- vbagr10.chm65727
ms.prod: excel
api_name:
- Excel.Forward
ms.assetid: 6a2e78d9-12ca-160a-7154-4968054f6b72
ms.date: 06/08/2017
---


# Forward Property

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write  **Long**.


## Example

This example sets the number of units that the trendline extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```



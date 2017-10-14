---
title: HasErrorBars Property
keywords: vbagr10.chm65696
f1_keywords:
- vbagr10.chm65696
ms.prod: excel
api_name:
- Excel.HasErrorBars
ms.assetid: f16a2ffe-b481-ec32-1144-8c1e5718243f
ms.date: 06/08/2017
---


# HasErrorBars Property

 **True** if the series has error bars. This property isn't available for 3-D charts. Read/write **Boolean**.


## Example

This example removes error bars from series one. The example should be run on a 2-D line chart that has error bars for series one.


```vb
myChart.SeriesCollection(1).HasErrorBars = False
```



---
title: VaryByCategories Property
keywords: vbagr10.chm65596
f1_keywords:
- vbagr10.chm65596
ms.prod: excel
api_name:
- Excel.VaryByCategories
ms.assetid: e64bd5cb-1dfa-b78a-ee7e-cf3eb7b4a788
ms.date: 06/08/2017
---


# VaryByCategories Property

 **True** if Microsoft Graph assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write **Boolean**.


## Example

This example assigns a different color or pattern to each data marker in chart group one. The example should be run on a 2-D line chart that has data markers on a series.


```vb
myChart.ChartGroups(1).VaryByCategories = True
```



---
title: Overlap Property
keywords: vbagr10.chm5207749
f1_keywords:
- vbagr10.chm5207749
ms.prod: excel
api_name:
- Excel.Overlap
ms.assetid: 60e82754-4553-7ee9-7403-06cd12de733e
ms.date: 06/08/2017
---


# Overlap Property

Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts. Read/write  **Long**.


## Remarks

If this property is set to -100, bars are positioned so that there's one bar width between them. If the overlap is 0 (zero), there's no space between bars (one bar starts immediately after the preceding bar). If the overlap is 100, bars are positioned on top of each other.


## Example

This example sets the overlap for chart group one to -50. The example should be run on a 2-D column chart that has two or more series.


```
myChart.ChartGroups(1).Overlap = -50
```



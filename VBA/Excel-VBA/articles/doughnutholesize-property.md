---
title: DoughnutHoleSize Property
keywords: vbagr10.chm66662
f1_keywords:
- vbagr10.chm66662
ms.prod: excel
api_name:
- Excel.DoughnutHoleSize
ms.assetid: 07e1e63b-8e31-92e5-18ab-c47104d093ac
ms.date: 06/08/2017
---


# DoughnutHoleSize Property

Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. Read/write  **Long**.


## Example

This example sets the hole size for doughnut group one. The example should be run on a 2-D doughnut chart.


```
myChart.DoughnutGroups(1).DoughnutHoleSize = 10
```



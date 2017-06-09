---
title: FirstSliceAngle Property
keywords: vbagr10.chm65599
f1_keywords:
- vbagr10.chm65599
ms.prod: excel
api_name:
- Excel.FirstSliceAngle
ms.assetid: 53f1fa5e-71d5-bf71-0fec-5f7be85b02d2
ms.date: 06/08/2017
---


# FirstSliceAngle Property

Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Read/write  **Long**.


## Example

This example sets the angle for the first slice in chart group one. The example should be run on a 2-D pie chart.


```
myChart.ChartGroups(1).FirstSliceAngle = 15
```



---
title: Perspective Property
keywords: vbagr10.chm65593
f1_keywords:
- vbagr10.chm65593
ms.prod: excel
api_name:
- Excel.Perspective
ms.assetid: 84ddaf6c-1204-1a7b-55e5-7d3cf2787a2c
ms.date: 06/08/2017
---


# Perspective Property

Returns or sets the perspective for the 3-D chart view. Must be from 0 through 100. This property is ignored if the  **[RightAngleAxes](rightangleaxes-property.md)** property is **True**. Read/write  **Long**.


## Example

This example sets the perspective of  `myChart` to 70. The example should be run on a 3-D chart.


```vb
myChart.RightAngleAxes = False 
myChart.Perspective = 70
```



---
title: Smooth Property
keywords: vbagr10.chm65699
f1_keywords:
- vbagr10.chm65699
ms.prod: excel
api_name:
- Excel.Smooth
ms.assetid: 037fa5ed-dd47-c544-50c4-813bc8000955
ms.date: 06/08/2017
---


# Smooth Property

True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write Boolean.

 _expression_. **Smooth**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example turns on curve smoothing for series one. The example should be run on a 2-D line chart.


```vb
myChart.SeriesCollection(1).Smooth = True
```



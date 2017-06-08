---
title: LegendKey Property
keywords: vbagr10.chm65710
f1_keywords:
- vbagr10.chm65710
ms.prod: excel
api_name:
- Excel.LegendKey
ms.assetid: 55277508-2a81-c9c0-1f34-4d44c967ae8e
ms.date: 06/08/2017
---


# LegendKey Property

Returns a  **[LegendKey](legendkey-object.md)** object that represents the legend key associated with the entry.


## Example

This example sets the legend key for legend entry one to be a triangle. The example should be run on a 2-D line chart.


```
myChart.Legend.LegendEntries(1).LegendKey _ 
 .MarkerStyle = xlMarkerStyleTriangle
```



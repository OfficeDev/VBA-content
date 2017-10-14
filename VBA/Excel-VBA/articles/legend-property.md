---
title: Legend Property
keywords: vbagr10.chm5207602
f1_keywords:
- vbagr10.chm5207602
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: 03d13546-c567-04b3-8ed5-cb99dc97c8e4
ms.date: 06/08/2017
---


# Legend Property

Returns a  **[Legend](legend-object.md)** object that represents the legend for the specified chart. Read-only.


## Example

This example turns on the legend for the chart and then sets the font color for the legend to blue.


```vb
myChart.HasLegend = True 
myChart.Legend.Font.ColorIndex = 5
```



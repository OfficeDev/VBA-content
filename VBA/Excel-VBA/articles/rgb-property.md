---
title: RGB Property
keywords: vbagr10.chm5207930
f1_keywords:
- vbagr10.chm5207930
ms.prod: excel
api_name:
- Excel.RGB
ms.assetid: bb3dbad0-a96a-969d-1234-ee9cf59e4c87
ms.date: 06/08/2017
---


# RGB Property

Returns the red-green-blue value of the specified color. Read-only  **Long**.


## Example

This example sets the color of the legend font to the foreground fill color of the plot area.


```
myChart.Legend.Font.Color = _ 
 myChart.PlotArea.Fill.ForeColor.RGB
```



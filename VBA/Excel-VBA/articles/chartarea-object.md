---
title: ChartArea Object
keywords: vbagr10.chm5207179
f1_keywords:
- vbagr10.chm5207179
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 85fcf460-6b2b-142f-ce4a-4a74e9d8efd3
ms.date: 06/08/2017
---


# ChartArea Object

Represents the chart area of the specified chart. The chart area in a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area in a 3-D chart contains the chart title and the legend; it doesn't include the plot area (the area within the chart area where the data is plotted). For information about formatting the plot area, see the  **[PlotArea](plotarea-object.md)** object.


## Using the ChartArea Object

Use the  **ChartArea** property to return the **ChartArea** object. The following example sets the pattern for the chart area.


```
myChart.ChartArea.Interior.Pattern = xlLightDown
```



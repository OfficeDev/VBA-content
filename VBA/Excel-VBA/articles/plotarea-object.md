---
title: PlotArea Object
keywords: vbagr10.chm131210
f1_keywords:
- vbagr10.chm131210
ms.prod: excel
api_name:
- Excel.PlotArea
ms.assetid: 49763ddd-3039-d15c-4ec4-e3b4f4e08d84
ms.date: 06/08/2017
---


# PlotArea Object

Represents the plot area of the specified chart. This is the area where your chart data is plotted. The plot area in a 2-D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area in a 3-D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area in a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area in a 3-D chart contains the chart title and the legend. For information about formatting the chart area, see the  **[ChartArea](chartarea-object.md)** object.

## Using the PlotArea Object

Use the  **PlotArea** property to return the **PlotArea** object. The following example places a dashed border around the chart area and places a dotted border around the plot area.


```vb
With myChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```



---
title: ChartTitle.IncludeInLayout Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartTitle.IncludeInLayout
ms.assetid: d4942d3e-1c58-c3b5-c291-64bf64300f9e
ms.date: 06/08/2017
---


# ChartTitle.IncludeInLayout Property (PowerPoint)

 **True** if a chart title will occupy the chart layout space when a chart layout is being determined. The default is **True**. Read/write **Boolean**.


## Syntax

 _expression_. **IncludeInLayout**

 _expression_ An expression that returns a **[ChartTitle](charttitle-object-powerpoint.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[ChartTitle Object](charttitle-object-powerpoint.md)


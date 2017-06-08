---
title: Legend.IncludeInLayout Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Legend.IncludeInLayout
ms.assetid: 2e14a6e0-923b-d383-2e40-dfa17f95df92
ms.date: 06/08/2017
---


# Legend.IncludeInLayout Property (PowerPoint)

 **True** if a legend will occupy the chart layout space when a chart layout is being determined. The default is **True**. Read/write **Boolean**.


## Syntax

 _expression_. **IncludeInLayout**

 _expression_ A variable that represents a **[Legend](legend-object-powerpoint.md)** object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title by using the  **Above Chart** command, the chart will resize smaller, as in previous versions of Microsoft Office. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


#### Concepts


[Legend Object](legend-object-powerpoint.md)


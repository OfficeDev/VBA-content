---
title: Walls Property
keywords: vbagr10.chm65622
f1_keywords:
- vbagr10.chm65622
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 74da4bfa-7b53-80d9-a673-42a67ffab787
ms.date: 06/08/2017
---


# Walls Property

Returns a  **[Walls](walls-object.md)** collection that represents the walls of the 3-D chart. Read-only.


## Remarks

This property doesn't apply to 3-D pie charts.


## Example

This example sets the color of the wall border of the chart to red. The example should be run on a 3-D chart.


```
myChart.Walls.Border.ColorIndex = 3
```



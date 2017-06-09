---
title: DropLines Object
keywords: vbagr10.chm5207329
f1_keywords:
- vbagr10.chm5207329
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 52fa64aa-0b0b-bbe1-1ec2-d866e2e35674
ms.date: 06/08/2017
---


# DropLines Object

Represents the drop lines in the specified chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object isn't a collection. There's no object that represents a single drop line; either you have drop lines turned on for all points in a chart group or you have them turned off.


## Using the DropLines Object

Use the  **DropLines** property to return the **DropLines** object. The following example turns on drop lines for chart group one in the chart and then sets the drop-line color to red.


```vb
myChart.ChartGroups(1).HasDropLines = True 
myChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
```


## Remarks

If the  **[HasDropLines](hasdroplines-property.md)** property is  **False**, most properties of the  **DropLines** object are disabled.



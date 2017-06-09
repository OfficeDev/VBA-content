---
title: Border Object
keywords: vbagr10.chm5207143
f1_keywords:
- vbagr10.chm5207143
ms.prod: excel
api_name:
- Excel.Border
ms.assetid: cb5ee6ef-f497-5113-85e4-a312871ad072
ms.date: 06/08/2017
---


# Border Object

Represents the border of the specified object.


## Using the Border Object

An object's border is treated as a single entity and is always returned as a unit (in its entirety), regardless of how many sides it has. Use the  **Border** property to return the **Border** object. The following example places a dashed border around the chart area and places a dotted border around the plot area.


```vb
With myChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```



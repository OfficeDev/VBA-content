---
title: TickLabels Property
keywords: vbagr10.chm65627
f1_keywords:
- vbagr10.chm65627
ms.prod: excel
api_name:
- Excel.TickLabels
ms.assetid: 5aa48053-c9ff-71c7-7a03-d7fe47e681c7
ms.date: 06/08/2017
---


# TickLabels Property

Returns a  **[TickLabels](ticklabels-object.md)** collection that represents the tick-mark labels for the specified axis. Read-only.


## Example

This example sets the color of the tick-mark label font for the value axis.


```
myChart.Axes(xlValue).TickLabels.Font.ColorIndex = 3
```



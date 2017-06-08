---
title: TickLabels Object
keywords: vbagr10.chm5208056
f1_keywords:
- vbagr10.chm5208056
ms.prod: excel
api_name:
- Excel.TickLabels
ms.assetid: d71b6cf2-c4ad-66f3-f7c2-8219f9ec21b1
ms.date: 06/08/2017
---


# TickLabels Object

Represents the tick-mark labels associated with tick marks on the specified chart axis. This object isn't a collection. There's no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the  **TickLabelSpacing** property for the category axis.

Tick-mark label text for the value axis is calculated based on the  **MajorUnit**,  **MinimumScale**, and  **MaximumScale** properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.


## Using the TickLabels Object

Use the  **TickLabels** property to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis in the chart.


```
myChart.Axes(xlValue).TickLabels.NumberFormat = "0.00"
```



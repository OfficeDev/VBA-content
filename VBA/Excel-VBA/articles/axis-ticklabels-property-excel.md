---
title: Axis.TickLabels Property (Excel)
keywords: vbaxl10.chm561100
f1_keywords:
- vbaxl10.chm561100
ms.prod: excel
api_name:
- Excel.Axis.TickLabels
ms.assetid: e8a6edf9-2fdd-d8e9-0de9-5c4aa921c6b1
ms.date: 06/08/2017
---


# Axis.TickLabels Property (Excel)

Returns a  **[TickLabels](ticklabels-object-excel.md)** object that represents the tick-mark labels for the specified axis. Read-only.


## Syntax

 _expression_ . **TickLabels**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the color of the tick-mark label font for the value axis in Chart1.


```vb
Charts("Chart1").Axes(xlValue).TickLabels.Font.ColorIndex = 3
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)


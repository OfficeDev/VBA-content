---
title: TickLabelSpacing Property
keywords: vbagr10.chm5208063
f1_keywords:
- vbagr10.chm5208063
ms.prod: excel
api_name:
- Excel.TickLabelSpacing
ms.assetid: f8bf4611-3b25-3d66-f49b-5a088e95028b
ms.date: 06/08/2017
---


# TickLabelSpacing Property

Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Read/write  **Long**.


## Remarks

Tick-mark label spacing on the value axis is always calculated by Microsoft Graph.


## Example

This example sets the number of categories between tick-mark labels on the category axis.


```
myChart.Axes(xlCategory).TickLabelSpacing = 10
```



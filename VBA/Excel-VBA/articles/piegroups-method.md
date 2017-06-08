---
title: PieGroups Method
keywords: vbagr10.chm65549
f1_keywords:
- vbagr10.chm65549
ms.prod: excel
api_name:
- Excel.PieGroups
ms.assetid: f7fd5497-f7a0-6c28-1a59-9e6f37a0885e
ms.date: 06/08/2017
---


# PieGroups Method

On a 2-D chart, returns an object that represents either a single pie chart group or a collection of the pie chart groups.

 _expression_. **PieGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. Specifies the chart group.

## Example

This example sets pie group one to use a different color for each data marker. The example should be run on a 2-D chart.


```vb
myChart.PieGroups(1).VaryByCategories = True
```



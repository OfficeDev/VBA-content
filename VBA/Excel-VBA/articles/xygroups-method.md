---
title: XYGroups Method
keywords: vbagr10.chm3077639
f1_keywords:
- vbagr10.chm3077639
ms.prod: excel
api_name:
- Excel.XYGroups
ms.assetid: d334382a-8d27-2b35-4306-a16f5fa13c89
ms.date: 06/08/2017
---


# XYGroups Method

On a 2-D chart, returns an object that represents either a single scatter chart group or a collection of the scatter chart groups.

 _expression_. **XYGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. Specifies the chart group.

## Example

This example sets X-Y group (scatter group) one to use a different color for each data marker. The example should be run on a 2-D chart.


```vb
myChart.XYGroups(1).VaryByCategories = True
```



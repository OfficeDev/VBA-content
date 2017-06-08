---
title: BarGroups Method
keywords: vbagr10.chm65546
f1_keywords:
- vbagr10.chm65546
ms.prod: excel
api_name:
- Excel.BarGroups
ms.assetid: a00e484e-05ec-2eaa-cc33-05b77a4af0b5
ms.date: 06/08/2017
---


# BarGroups Method

On a 2-D chart, this method returns an object that represents either a single bar chart group or a collection of all the bar chart groups.

 _expression_. **BarGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The index number of the specified bar chart group.

## Example

This example sets the space between bar clusters in the 2-D bar chart group to be 50 percent of the bar width.


```
myChart.BarGroups(1).GapWidth = 50
```



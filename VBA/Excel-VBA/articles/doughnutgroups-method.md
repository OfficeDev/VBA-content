---
title: DoughnutGroups Method
keywords: vbagr10.chm3077618
f1_keywords:
- vbagr10.chm3077618
ms.prod: excel
api_name:
- Excel.DoughnutGroups
ms.assetid: 41ca4213-c17b-7bba-c357-7ba65fd55d39
ms.date: 06/08/2017
---


# DoughnutGroups Method

On a 2-D chart, returns an object that represents either a single doughnut chart group or a collection of the doughnut chart groups.

 _expression_. **DoughnutGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. Specifies the chart group.

## Example

This example sets the starting angle for doughnut group one.


```
myChart.DoughnutGroups(1).FirstSliceAngle = 45
```



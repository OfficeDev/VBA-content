---
title: LineGroups Method
keywords: vbagr10.chm3077623
f1_keywords:
- vbagr10.chm3077623
ms.prod: excel
api_name:
- Excel.LineGroups
ms.assetid: 3a8083b5-8b71-e28b-c775-6be50544d6b2
ms.date: 06/08/2017
---


# LineGroups Method

On a 2-D chart, returns an object that represents either a single line chart group or a collection of the line chart groups.

 _expression_. **LineGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. Specifies the chart group.

## Example

This example sets line group one to use a different color for each data marker. The example should be run on a 2-D chart.


```vb
myChart.LineGroups(1).VaryByCategories = True
```



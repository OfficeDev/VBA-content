---
title: RadarGroups Method
keywords: vbagr10.chm3077630
f1_keywords:
- vbagr10.chm3077630
ms.prod: excel
api_name:
- Excel.RadarGroups
ms.assetid: 5fbca532-ae99-fb41-dd1d-2d24909bf073
ms.date: 06/08/2017
---


# RadarGroups Method

On a 2-D chart, returns an object that represents either a single radar chart group or a collection of the radar chart groups.

 _expression_. **RadarGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. Specifies the chart group.

## Example

This example sets radar group one to use a different color for each data marker. The example should be run on a 2-D chart.


```vb
myChart.RadarGroups(1).VaryByCategories = True
```



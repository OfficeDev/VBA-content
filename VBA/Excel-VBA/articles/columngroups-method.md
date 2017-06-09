---
title: ColumnGroups Method
keywords: vbagr10.chm65547
f1_keywords:
- vbagr10.chm65547
ms.prod: excel
api_name:
- Excel.ColumnGroups
ms.assetid: dcb4d7e0-ce56-46d9-35d9-d9653bbb6f97
ms.date: 06/08/2017
---


# ColumnGroups Method

On a 2-D chart, returns an object that represents either a single column chart group or a collection of the column chart groups.

 _expression_. **ColumnGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The index number of the specified column chart group.

## Example

This example sets the space between column clusters in the 2-D column chart group to be 50 percent of the column width.


```
myChart.ColumnGroups(1).GapWidth = 50
```



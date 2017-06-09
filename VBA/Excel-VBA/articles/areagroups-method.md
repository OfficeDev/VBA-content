---
title: AreaGroups Method
keywords: vbagr10.chm3077607
f1_keywords:
- vbagr10.chm3077607
ms.prod: excel
api_name:
- Excel.AreaGroups
ms.assetid: ec2a4a28-2f10-4f4f-bd91-642bf1b8ebe2
ms.date: 06/08/2017
---


# AreaGroups Method

On a 2-D chart, this method returns an object that represents a single area chart group or a collection of all the area chart groups.

 _expression_. **AreaGroups**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The index number of the specified chart group.

## Example

This example turns on drop lines for the 2-D area chart group.


```vb
myChart.AreaGroups(1).HasDropLines = True
```



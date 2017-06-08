---
title: Column3DGroup Property
keywords: vbagr10.chm3076976
f1_keywords:
- vbagr10.chm3076976
ms.prod: excel
api_name:
- Excel.Column3DGroup
ms.assetid: 9fa90f46-29b8-c710-93de-4150e276330c
ms.date: 06/08/2017
---


# Column3DGroup Property

Returns a ChartGroup object that represents the specified column chart group on a 3-D chart. Read-only ChartGroup object.

 _expression_. **Column3DGroup**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example sets the space between column clusters in the 3-D column chart group to be 50 percent of the column width.


```
myChart.Column3DGroup.GapWidth = 50
```



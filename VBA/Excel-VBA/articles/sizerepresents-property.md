---
title: SizeRepresents Property
keywords: vbagr10.chm67188
f1_keywords:
- vbagr10.chm67188
ms.prod: excel
api_name:
- Excel.SizeRepresents
ms.assetid: 54f87d5a-e388-e1d1-8a20-bec820f3449c
ms.date: 06/08/2017
---


# SizeRepresents Property

Returns or sets what the bubble size represents on a bubble chart. Read/write XlSizeRepresents .



|XlSizeRepresents can be one of these XlSizeRepresents constants.|
| **xlSizeIsArea**|
| **xlSizeIsWidth**|

 _expression_. **SizeRepresents**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets what the bubble size represents for the chart. (The example assumes that the chart is a bubble chart.)


```
myChart.ChartGroups(1).SizeRepresents = xlSizeIsWidth
```



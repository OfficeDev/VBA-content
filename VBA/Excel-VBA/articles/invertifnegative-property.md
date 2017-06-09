---
title: InvertIfNegative Property
keywords: vbagr10.chm65668
f1_keywords:
- vbagr10.chm65668
ms.prod: excel
api_name:
- Excel.InvertIfNegative
ms.assetid: 0b75c2af-85f5-86bb-ab7e-3eed3f88940e
ms.date: 06/08/2017
---


# InvertIfNegative Property

True if Microsoft Graph inverts the pattern in the item when it corresponds to a negative number. Read/write Boolean for all objects, except for the Interior object, which is read/write Variant.

 _expression_. **InvertIfNegative**

 _expression_ Required. An expression that returns one of the above objects.


## Example

This example inverts the pattern for negative values in series one. The example should be run on a 2-D column chart.


```vb
myChart.SeriesCollection(1).InvertIfNegative = True
```



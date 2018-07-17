---
title: Name Property (Graph)
keywords: vbagr10.chm3077554
f1_keywords:
- vbagr10.chm3077554
ms.prod: excel
ms.assetid: d3590902-6957-8e32-e627-5946ba66c44f
ms.date: 06/08/2017
---


# Name Property (Graph)

Name property as it applies to the  **Application** and **Trendline** objects.

Returns or sets the name of the object. Read/write String.

 _expression_. **Name**

 _expression_ Required. An expression that returns one of the above objects.
Name property as it applies to the  **Font** object.
Returns or sets the name of the object. Read/write Variant.
 _expression_. **Name**
 _expression_ Required. An expression that returns a **Font** object.
Name property as it applies to the all other objects.
Returns or sets the name of the object. Read-only String.
 _expression_. **Name**
 _expression_ Required. An expression that returns one of the above objects.

## Example

This example assigns the name of the first trendline to the variable myTrendname.


```
myTrendname = myChart.SeriesCollection(1).Trendlines(1).Name
```



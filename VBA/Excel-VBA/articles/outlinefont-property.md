---
title: OutlineFont Property
keywords: vbagr10.chm65757
f1_keywords:
- vbagr10.chm65757
ms.prod: excel
api_name:
- Excel.OutlineFont
ms.assetid: 41075763-8ee7-e6ba-c9a2-7bc718b5405e
ms.date: 06/08/2017
---


# OutlineFont Property

True if the font is an outline font. Read/write Variant.

 _expression_. **OutlineFont**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

This property has no effect in Windows, but its value is retained (it can be set and returned).


## Example

This example sets the font for the chart title to an outline font.


```vb
myChart.ChartTitle.Font.OutlineFont = True
```



---
title: LegendEntries Method
keywords: vbagr10.chm3077622
f1_keywords:
- vbagr10.chm3077622
ms.prod: excel
api_name:
- Excel.LegendEntries
ms.assetid: 6419aa89-6e59-dc04-ab79-67a0a38cad6f
ms.date: 06/08/2017
---


# LegendEntries Method

Returns an object that represents either a single legend entry or a collection of legend entries for the legend.

 _expression_. **LegendEntries**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The number of the legend entry.

## Example

This example sets the font for legend entry one.


```
myChart.Legend.LegendEntries(1).Font.Name = "Arial"
```



---
title: MarkerBackgroundColorIndex Property
keywords: vbagr10.chm65610
f1_keywords:
- vbagr10.chm65610
ms.prod: excel
api_name:
- Excel.MarkerBackgroundColorIndex
ms.assetid: 97995a64-0c94-3c55-ba73-9b5dedda4f2c
ms.date: 06/08/2017
---


# MarkerBackgroundColorIndex Property

Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants. Read/write XlColorIndex .



|XlColorIndex can be one of these XlColorIndex constants.|
| **xlColorIndexNone**|
| **xlColorIndexAutomatic**|

 _expression_. **MarkerBackgroundColorIndex**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the marker background and foreground colors for the second point in series one.


```vb
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColorIndex = 4 'green 
 .MarkerForegroundColorIndex = 3 'red 
End With
```



---
title: PatternColorIndex Property
keywords: vbagr10.chm3077568
f1_keywords:
- vbagr10.chm3077568
ms.prod: excel
api_name:
- Excel.PatternColorIndex
ms.assetid: d11aa18c-b46d-950c-78ef-e58dd1c751fb
ms.date: 06/08/2017
---


# PatternColorIndex Property

Returns or sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants. Read/write Variant.



|XlColorIndex can be one of these XlColorIndex constants.|
| **xlColorIndexAutomatic**|
| **xlColorIndexNone**|

 _expression_. **PatternColorIndex**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Remarks

Set this property to  **xlColorIndexAutomatic** to specify the automatic pattern for cells or the automatic fill style for drawing objects. Set this property to **xlColorIndexNone** to specify that you don't want a pattern (this is the same as setting the **Pattern** property of the **Interior** object to **xlPatternNone**).


## Example

This example sets the color of the interior pattern for the chart area.


```
myChart.ChartArea.Interior.PatternColorIndex = 5
```



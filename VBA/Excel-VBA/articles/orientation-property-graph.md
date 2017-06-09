---
title: Orientation Property (Graph)
keywords: vbagr10.chm65670
f1_keywords:
- vbagr10.chm65670
ms.prod: excel
ms.assetid: 1e4e111c-5144-a509-4791-e8ca31c3de5e
ms.date: 06/08/2017
---


# Orientation Property (Graph)

Returns or sets the text orientation. Can be an integer value from -90 degrees to 90 degrees or one of the following XlOrientation constants. Read/write XlTickLabelOrientation for all objects, except for the TickLabels object, which is read/write Variant.



|XlTickLabelOrientation can be one of these XlTickLabelOrientation constants.|
| **xlTickLabelOrientationAutomatic**|
| **xlTickLabelOrientationDownward**|
| **xlTickLabelOrientationHorizontal**|
| **xlTickLabelOrientationUpward**|
| **xlTickLabelOrientationVertical**|

 _expression_. **Orientation**

 _expression_ Required. An expression that returns one of the above objects.

## Example

This example sets the orientation for the chart title.


```
myChart.ChartTitle.Orientation = xlHorizontal
```



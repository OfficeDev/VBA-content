---
title: Position Property
keywords: vbagr10.chm65669
f1_keywords:
- vbagr10.chm65669
ms.prod: excel
api_name:
- Excel.Position
ms.assetid: 0e9e41e2-30a8-c744-72d1-3820cc4975f2
ms.date: 06/08/2017
---


# Position Property

Position property as it applies to the  **DataLabel** and **DataLabels** objects.

Returns or sets the position of the data label. Read/write XlDataLabelPosition.


|XlDataLabelPosition can be one of these XlDataLabelPosition constants.|
| **xlLabelPositionBelow**|
| **xlLabelPositionCenter**|
| **xlLabelPositionInsideBase**|
| **xlLabelPositionInsideEnd**|
| **xlLabelPositionLeft**|
| **xlLabelPositionMixed**|
| **xlLabelPositionOutsideEnd**|
| **xlLabelPositionRight**|
| **xlLabelPositionAbove**|
| **xlLabelPositionBestFit**|
| **xlLabelPositionCustom**|
 _expression_. **Position**
 _expression_ Required. An expression that returns one of the above objects.
Position property as it applies to the  **Legend** object.
Returns or sets the position of the legend on the chart. Read/write XlLegendPosition .


|XlLegendPosition can be one of these XlLegendPosition constants.|
| **xlLegendPositionBottom**|
| **xlLegendPositionCorner**|
| **xlLegendPositionLeft**|
| **xlLegendPositionRight**|
| **xlLegendPositionTop**|
 _expression_. **Position**
 _expression_ Required. An expression that returns one of the above objects.

## Example

This example sets the position of the legend to the top of the chart.


```
myChart.Legend.Position = xlLegendPositionTop
```



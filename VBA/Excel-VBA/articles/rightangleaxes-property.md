---
title: RightAngleAxes Property
keywords: vbagr10.chm3077581
f1_keywords:
- vbagr10.chm3077581
ms.prod: excel
api_name:
- Excel.RightAngleAxes
ms.assetid: 5c34e5b4-a936-70a5-cd0c-d9a7a091e8d0
ms.date: 06/08/2017
---


# RightAngleAxes Property

True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts. Read/write Variant.

 _expression_. **RightAngleAxes**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

If this property is  **True**, the  **[Perspective](perspective-property.md)** property is ignored.


## Example

This example sets the axes to intersect at right angles. The example should be run on a 3-D chart.


```vb
myChart.RightAngleAxes = True
```



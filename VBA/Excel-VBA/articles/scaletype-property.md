---
title: ScaleType Property
keywords: vbagr10.chm3077583
f1_keywords:
- vbagr10.chm3077583
ms.prod: excel
api_name:
- Excel.ScaleType
ms.assetid: 500fa5e4-4e19-bdd4-fa28-4dcba763c8a7
ms.date: 06/08/2017
---


# ScaleType Property

Returns or sets the value axis scale type. Applies only to the value axis. Read/write XlScaleType .



|XlScaleType can be one of these XlScaleType constants.|
| **xlScaleLinear**|
| **xlScaleLogarithmic**|

 _expression_. **ScaleType**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Remarks

A logarithmic scale uses base 10 logarithms.


## Example

This example sets the value axis to use a logarithmic scale.


```
myChart.Axes(xlValue).ScaleType = xlScaleLogarithmic
```



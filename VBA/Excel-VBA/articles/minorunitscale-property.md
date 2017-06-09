---
title: MinorUnitScale Property
keywords: vbagr10.chm67186
f1_keywords:
- vbagr10.chm67186
ms.prod: excel
api_name:
- Excel.MinorUnitScale
ms.assetid: c246ab1e-5c41-f15e-fdbc-d219f2d03448
ms.date: 06/08/2017
---


# MinorUnitScale Property

Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit .



|XlTimeUnit can be one of these XlTimeUnit constants.|
| **xlDays**|
| **xlMonths**|
| **xlYears**|

 _expression_. **MinorUnitScale**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the category axis to use a time scale and sets the major and minor units.


```vb
With myChart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .MajorUnit = 5 
 .MajorUnitScale = xlDays 
 .MinorUnit = 1 
 .MinorUnitScale = xlDays 
End With
```



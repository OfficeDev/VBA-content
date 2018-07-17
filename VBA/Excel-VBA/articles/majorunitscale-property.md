---
title: MajorUnitScale Property
keywords: vbagr10.chm67185
f1_keywords:
- vbagr10.chm67185
ms.prod: excel
api_name:
- Excel.MajorUnitScale
ms.assetid: b2a54ca7-6eac-5552-6de7-ee0ab59e1ddb
ms.date: 06/08/2017
---


# MajorUnitScale Property

Returns or sets the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit .



|XlTimeUnit can be one of these XlTimeUnit constants.|
| **xlDays**|
| **xlMonths**|
| **xlYears**|

 _expression_. **MajorUnitScale**

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



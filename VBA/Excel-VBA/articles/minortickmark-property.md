---
title: MinorTickMark Property
keywords: vbagr10.chm65563
f1_keywords:
- vbagr10.chm65563
ms.prod: excel
api_name:
- Excel.MinorTickMark
ms.assetid: cbb515d8-fdae-2546-f13b-80ed75cc4192
ms.date: 06/08/2017
---


# MinorTickMark Property

Returns or sets the type of minor tick mark for the specified axis. Read/write XlTickMark .



|XlTickMark can be one of these XlTickMark constants.|
| **xlTickMarkCross**|
| **xlTickMarkInside**|
| **xlTickMarkNone**|
| **xlTickMarkOutside**|

 _expression_. **MinorTickMark**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the minor tick marks for the value axis to be inside the axis.


```
myChart.Axes(xlValue).MinorTickMark = xlTickMarkInside
```



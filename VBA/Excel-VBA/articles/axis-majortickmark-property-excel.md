---
title: Axis.MajorTickMark Property (Excel)
keywords: vbaxl10.chm561085
f1_keywords:
- vbaxl10.chm561085
ms.prod: excel
api_name:
- Excel.Axis.MajorTickMark
ms.assetid: 0b481503-76a8-2b04-8c61-0fef649ce03e
ms.date: 06/08/2017
---


# Axis.MajorTickMark Property (Excel)

Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-excel.md)** .


## Syntax

 _expression_ . **MajorTickMark**

 _expression_ A variable that represents an **Axis** object.


## Remarks





| **XlTickMark** can be one of these **XlTickMark** constants.|
| **xlTickMarkInside**|
| **xlTickMarkOutside**|
| **xlTickMarkCross**|
| **xlTickMarkNone**|

## Example

This example sets the major tick marks for the value axis in Chart1 to be outside the axis.


```vb
Charts("Chart1").Axes(xlValue).MajorTickMark = xlTickMarkOutside
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)


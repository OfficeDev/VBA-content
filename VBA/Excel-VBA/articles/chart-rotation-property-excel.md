---
title: Chart.Rotation Property (Excel)
keywords: vbaxl10.chm149139
f1_keywords:
- vbaxl10.chm149139
ms.prod: excel
api_name:
- Excel.Chart.Rotation
ms.assetid: bf271f86-18c9-ac74-12ab-f90f4353f71d
ms.date: 06/08/2017
---


# Chart.Rotation Property (Excel)

Returns or sets the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write  **Variant** .


## Syntax

 _expression_ . **Rotation**

 _expression_ A variable that represents a **Chart** object.


## Remarks

Rotations are always rounded to the nearest integer.


## Example

This example sets the rotation of Chart1 to 30 degrees. The example should be run on a 3-D chart.


```vb
Charts("Chart1").Rotation = 30
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)


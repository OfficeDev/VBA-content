---
title: ChartGroup.FirstSliceAngle Property (Excel)
keywords: vbaxl10.chm568077
f1_keywords:
- vbaxl10.chm568077
ms.prod: excel
api_name:
- Excel.ChartGroup.FirstSliceAngle
ms.assetid: a6bded62-d757-fc67-4677-7f9c12fd6395
ms.date: 06/08/2017
---


# ChartGroup.FirstSliceAngle Property (Excel)

Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/write  **Long** .


## Syntax

 _expression_ . **FirstSliceAngle**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example sets the angle for the first slice in chart group one in Chart1. The example should be run on a 2-D pie chart.


```vb
Charts("Chart1").ChartGroups(1).FirstSliceAngle = 15
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


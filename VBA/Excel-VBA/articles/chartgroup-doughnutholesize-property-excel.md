---
title: ChartGroup.DoughnutHoleSize Property (Excel)
keywords: vbaxl10.chm568074
f1_keywords:
- vbaxl10.chm568074
ms.prod: excel
api_name:
- Excel.ChartGroup.DoughnutHoleSize
ms.assetid: d7c2cfbe-209b-2276-1245-2ddc31c1f93d
ms.date: 06/08/2017
---


# ChartGroup.DoughnutHoleSize Property (Excel)

Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. Read/write  **Long** .


## Syntax

 _expression_ . **DoughnutHoleSize**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example sets the hole size for doughnut group one in Chart1. The example should be run on a 2-D doughnut chart.


```vb
Charts("Chart1").DoughnutGroups(1).DoughnutHoleSize = 10
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


---
title: PlotBy Property
keywords: vbagr10.chm65738
f1_keywords:
- vbagr10.chm65738
ms.prod: excel
api_name:
- Excel.PlotBy
ms.assetid: 9cbc8692-0b50-1b46-c945-a3594a5d29b2
ms.date: 06/08/2017
---


# PlotBy Property

Returns or sets the way columns or rows are used as data series on the chart. Read/write XlRowCol.



|XlRowCol can be one of these XlRowCol constants.|
| **xlColumns**|
| **xlRows**|

 _expression_. **PlotBy**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example causes the embedded chart to plot data by columns.


```
myChart.PlotBy = xlColumns
```



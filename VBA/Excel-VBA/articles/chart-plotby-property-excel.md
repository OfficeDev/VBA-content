---
title: Chart.PlotBy Property (Excel)
keywords: vbaxl10.chm149155
f1_keywords:
- vbaxl10.chm149155
ms.prod: excel
api_name:
- Excel.Chart.PlotBy
ms.assetid: 69ff0fbe-7954-6808-68fa-cc92b2851dd8
ms.date: 06/08/2017
---


# Chart.PlotBy Property (Excel)

Returns or sets the way columns or rows are used as data series on the chart. Can be one of the following  **[XlRowCol](xlrowcol-enumeration-excel.md)** constants: **xlColumns** or **xlRows** . Read/write **Long** .


## Syntax

 _expression_ . **PlotBy**

 _expression_ A variable that represents a **Chart** object.


## Remarks

For PivotChart reports, this property is read-only and always returns  **xlColumns** .


## Example

This example causes the embedded chart to plot data by columns.


```vb
Worksheets(1).ChartObjects(1).Chart.PlotBy = xlColumns
```


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)
[Chart Object](chart-object-excel.md)


---
title: Workbook.ActiveChart Property (Excel)
keywords: vbaxl10.chm199075
f1_keywords:
- vbaxl10.chm199075
ms.prod: excel
api_name:
- Excel.Workbook.ActiveChart
ms.assetid: 81e18252-b1fe-2487-535e-6e24c80bef24
ms.date: 06/08/2017
---


# Workbook.ActiveChart Property (Excel)

Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .


## Syntax

 _expression_ . **ActiveChart**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

If you don't specify an object qualifier, this property returns the active chart in the active workbook.


## Example

This example turns on the legend for the active chart.


```vb
ActiveChart.HasLegend = True
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)


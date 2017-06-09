---
title: Application.ActiveChart Property (Excel)
keywords: vbaxl10.chm183075
f1_keywords:
- vbaxl10.chm183075
ms.prod: excel
api_name:
- Excel.Application.ActiveChart
ms.assetid: 37b1901c-a9c2-4a86-ce05-22f3989bc9d8
ms.date: 06/08/2017
---


# Application.ActiveChart Property (Excel)

Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .


## Syntax

 _expression_ . **ActiveChart**

 _expression_ A variable that represents an **Application** object.


## Remarks

If you don't specify an object qualifier, this property returns the active chart in the active workbook.


## Example

This example turns on the legend for the active chart.


```vb
ActiveChart.HasLegend = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)


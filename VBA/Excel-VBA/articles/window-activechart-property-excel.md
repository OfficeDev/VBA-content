---
title: Window.ActiveChart Property (Excel)
keywords: vbaxl10.chm356077
f1_keywords:
- vbaxl10.chm356077
ms.prod: excel
api_name:
- Excel.Window.ActiveChart
ms.assetid: 505902dd-63c3-cd11-c3cc-a82680c11768
ms.date: 06/08/2017
---


# Window.ActiveChart Property (Excel)

Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .


## Syntax

 _expression_ . **ActiveChart**

 _expression_ A variable that represents a **Window** object.


## Remarks

If you don't specify an object qualifier, this property returns the active chart in the active workbook.


## Example

This example turns on the legend for the active chart.


```vb
ActiveChart.HasLegend = True
```


## See also


#### Concepts


[Window Object](window-object-excel.md)


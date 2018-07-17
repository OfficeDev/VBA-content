---
title: Chart.Legend Property (Excel)
keywords: vbaxl10.chm149120
f1_keywords:
- vbaxl10.chm149120
ms.prod: excel
api_name:
- Excel.Chart.Legend
ms.assetid: 6396ca0f-63b5-3d4a-4f6b-b4e80a1911b3
ms.date: 06/08/2017
---


# Chart.Legend Property (Excel)

Returns a  **[Legend](legend-object-excel.md)** object that represents the legend for the chart. Read-only.


## Syntax

 _expression_ . **Legend**

 _expression_ A variable that represents a **Chart** object.


## Example

This example turns on the legend for Chart1 and then sets the legend font color to blue.


```vb
Charts("Chart1").HasLegend = True 
Charts("Chart1").Legend.Font.ColorIndex = 5
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)


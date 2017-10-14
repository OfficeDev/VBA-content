---
title: Chart.DisplayBlanksAs Property (Excel)
keywords: vbaxl10.chm149101
f1_keywords:
- vbaxl10.chm149101
ms.prod: excel
api_name:
- Excel.Chart.DisplayBlanksAs
ms.assetid: b4e18939-6214-25e8-a0cd-c984b9f82346
ms.date: 06/08/2017
---


# Chart.DisplayBlanksAs Property (Excel)

Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-excel.md)** constants. Read/write **Long** .


## Syntax

 _expression_ . **DisplayBlanksAs**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets Microsoft Excel to not plot blank cells in Chart1.


```vb
Charts("Chart1").DisplayBlanksAs = xlNotPlotted
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)
[SparklineGroup Object](sparklinegroup-object-excel.md)


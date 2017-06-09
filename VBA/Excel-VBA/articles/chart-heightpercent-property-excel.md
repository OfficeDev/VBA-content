---
title: Chart.HeightPercent Property (Excel)
keywords: vbaxl10.chm149117
f1_keywords:
- vbaxl10.chm149117
ms.prod: excel
api_name:
- Excel.Chart.HeightPercent
ms.assetid: a95f2b76-57a1-4c04-9f5f-ccd7852d4ab6
ms.date: 06/08/2017
---


# Chart.HeightPercent Property (Excel)

Returns or sets the height of a 3-D chart as a percentage of the chart width (between 5 and 500 percent). Read/write  **Long** .


## Syntax

 _expression_ . **HeightPercent**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the height of Chart1 to 80 percent of its width. The example should be run on a 3-D chart.


```vb
Charts("Chart1").HeightPercent = 80
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)


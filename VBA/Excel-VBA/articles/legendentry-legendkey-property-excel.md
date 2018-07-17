---
title: LegendEntry.LegendKey Property (Excel)
keywords: vbaxl10.chm586077
f1_keywords:
- vbaxl10.chm586077
ms.prod: excel
api_name:
- Excel.LegendEntry.LegendKey
ms.assetid: 727de973-636f-1018-5fc0-809a6af3a6f5
ms.date: 06/08/2017
---


# LegendEntry.LegendKey Property (Excel)

Returns a  **[LegendKey](legendkey-object-excel.md)** object that represents the legend key associated with the entry.


## Syntax

 _expression_ . **LegendKey**

 _expression_ A variable that represents a **LegendEntry** object.


## Example

This example sets the legend key for legend entry one on Chart1 to be a triangle. The example should be run on a 2-D line chart.


```vb
Charts("Chart1").Legend.LegendEntries(1).LegendKey _ 
 .MarkerStyle = xlMarkerStyleTriangle
```


## See also


#### Concepts


[LegendEntry Object](legendentry-object-excel.md)


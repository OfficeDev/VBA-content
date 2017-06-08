---
title: Series.Smooth Property (Excel)
keywords: vbaxl10.chm578106
f1_keywords:
- vbaxl10.chm578106
ms.prod: excel
api_name:
- Excel.Series.Smooth
ms.assetid: 24cb3fc6-a6ed-71ca-1aab-c1ea23480d00
ms.date: 06/08/2017
---


# Series.Smooth Property (Excel)

 **True** if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write.


## Syntax

 _expression_ . **Smooth**

 _expression_ A variable that represents a **Series** object.


## Example

This example turns on curve smoothing for series one in Chart1. The example should be run on a 2-D line chart.


```vb
Charts("Chart1").SeriesCollection(1).Smooth = True
```


## See also


#### Concepts


[Series Object](series-object-excel.md)


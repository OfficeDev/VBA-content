---
title: Point.SecondaryPlot Property (Excel)
keywords: vbaxl10.chm576098
f1_keywords:
- vbaxl10.chm576098
ms.prod: excel
api_name:
- Excel.Point.SecondaryPlot
ms.assetid: 1a12020a-bbd5-30b0-106a-589a44b45ca6
ms.date: 06/08/2017
---


# Point.SecondaryPlot Property (Excel)

 **True** if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts. Read/write **Boolean** .


## Syntax

 _expression_ . **SecondaryPlot**

 _expression_ A variable that represents a **Point** object.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example moves point four to the secondary section of the chart.


```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .Points(4).SecondaryPlot = True 
End With
```


## See also


#### Concepts


[Point Object](point-object-excel.md)


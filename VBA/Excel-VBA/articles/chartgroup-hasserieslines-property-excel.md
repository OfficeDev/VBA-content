---
title: ChartGroup.HasSeriesLines Property (Excel)
keywords: vbaxl10.chm568082
f1_keywords:
- vbaxl10.chm568082
ms.prod: excel
api_name:
- Excel.ChartGroup.HasSeriesLines
ms.assetid: 4285cf5b-ebb0-a6fd-49c1-d36c341bd016
ms.date: 06/08/2017
---


# ChartGroup.HasSeriesLines Property (Excel)

 **True** if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts. Read/write **Boolean** .


## Syntax

 _expression_ . **HasSeriesLines**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on series lines for chart group one in Chart1 and then sets their line style, weight, and color. The example should be run on a 2-D stacked column chart that has two or more series.


```vb
With Charts("Chart1").ChartGroups(1) 
 .HasSeriesLines = True 
 With .SeriesLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


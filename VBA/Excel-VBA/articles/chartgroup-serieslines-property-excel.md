---
title: ChartGroup.SeriesLines Property (Excel)
keywords: vbaxl10.chm568089
f1_keywords:
- vbaxl10.chm568089
ms.prod: excel
api_name:
- Excel.ChartGroup.SeriesLines
ms.assetid: 3e2156c3-c4dd-ef22-1645-ba27e7b499b8
ms.date: 06/08/2017
---


# ChartGroup.SeriesLines Property (Excel)

Returns a  **[SeriesLines](serieslines-object-excel.md)** object that represents the series lines for a 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie chart. Read-only.


## Syntax

 _expression_ . **SeriesLines**

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


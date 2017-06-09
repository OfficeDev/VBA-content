---
title: ChartGroup.SecondPlotSize Property (Excel)
keywords: vbaxl10.chm568099
f1_keywords:
- vbaxl10.chm568099
ms.prod: excel
api_name:
- Excel.ChartGroup.SecondPlotSize
ms.assetid: 231541fa-0353-3533-6b4b-0653b6157568
ms.date: 06/08/2017
---


# ChartGroup.SecondPlotSize Property (Excel)

Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/write  **Long** .


## Syntax

 _expression_ . **SecondPlotSize**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. The secondary section is 50 percent of the size of the primary pie.


```vb
With Worksheets(1).ChartObjects(1).Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
 .SecondPlotSize = 50 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


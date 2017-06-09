---
title: ChartGroup.UpBars Property (Excel)
keywords: vbaxl10.chm568092
f1_keywords:
- vbaxl10.chm568092
ms.prod: excel
api_name:
- Excel.ChartGroup.UpBars
ms.assetid: d97b23bd-4c51-2384-a5f3-7cc067d3d6fa
ms.date: 06/08/2017
---


# ChartGroup.UpBars Property (Excel)

Returns an  **[UpBars](upbars-object-excel.md)** object that represents the up bars on a line chart. Applies only to line charts. Read-only.


## Syntax

 _expression_ . **UpBars**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on up and down bars for chart group one in Chart1 and then sets their colors. The example should be run on a 2-D line chart containing two series that cross each other at one or more data points.


```vb
With Charts("Chart1").ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


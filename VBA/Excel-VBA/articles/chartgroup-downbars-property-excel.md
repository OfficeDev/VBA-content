---
title: ChartGroup.DownBars Property (Excel)
keywords: vbaxl10.chm568075
f1_keywords:
- vbaxl10.chm568075
ms.prod: excel
api_name:
- Excel.ChartGroup.DownBars
ms.assetid: dd8ae50c-0105-9645-467d-7eb07b56c95e
ms.date: 06/08/2017
---


# ChartGroup.DownBars Property (Excel)

Returns a  **[DownBars](downbars-object-excel.md)** object that represents the down bars on a line chart. Applies only to line charts. Read-only.


## Syntax

 _expression_ . **DownBars**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on up bars and down bars for chart group one in Chart1 and then sets their colors. The example should be run on a 2-D line chart that has two series that cross each other at one or more data points.


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


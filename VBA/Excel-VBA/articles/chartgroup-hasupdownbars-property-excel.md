---
title: ChartGroup.HasUpDownBars Property (Excel)
keywords: vbaxl10.chm568083
f1_keywords:
- vbaxl10.chm568083
ms.prod: excel
api_name:
- Excel.ChartGroup.HasUpDownBars
ms.assetid: 891f305c-521c-3ec5-3e88-886e1dbdaea2
ms.date: 06/08/2017
---


# ChartGroup.HasUpDownBars Property (Excel)

 **True** if a line chart has up and down bars. Applies only to line charts. Read/write **Boolean** .


## Syntax

 _expression_ . **HasUpDownBars**

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


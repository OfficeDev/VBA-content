---
title: ChartGroup.HiLoLines Property (Excel)
keywords: vbaxl10.chm568084
f1_keywords:
- vbaxl10.chm568084
ms.prod: excel
api_name:
- Excel.ChartGroup.HiLoLines
ms.assetid: 3d226065-9482-b393-a216-39d7c26961f0
ms.date: 06/08/2017
---


# ChartGroup.HiLoLines Property (Excel)

Returns a  **[HiLoLines](hilolines-object-excel.md)** object that represents the high-low lines for a series on a line chart. Applies only to line charts. Read-only.


## Syntax

 _expression_ . **HiLoLines**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on high-low lines for chart group one in Chart1 and then sets their line style, weight, and color. The example should be run on a 2-D line chart that has three series of stock-quote-like data (high-low-close).


```vb
With Charts("Chart1").ChartGroups(1) 
 .HasHiLoLines = True 
 With .HiLoLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


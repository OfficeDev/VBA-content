---
title: ChartGroup.DropLines Property (Excel)
keywords: vbaxl10.chm568076
f1_keywords:
- vbaxl10.chm568076
ms.prod: excel
api_name:
- Excel.ChartGroup.DropLines
ms.assetid: bf8ba5e6-ca8a-d664-f0b6-fe2af0353fbf
ms.date: 06/08/2017
---


# ChartGroup.DropLines Property (Excel)

Returns a  **[DropLines](droplines-object-excel.md)** object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.


## Syntax

 _expression_ . **DropLines**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on drop lines for chart group one in Chart1 and then sets their line style, weight, and color. The example should be run on a 2-D line chart that has one series.


```vb
With Charts("Chart1").ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


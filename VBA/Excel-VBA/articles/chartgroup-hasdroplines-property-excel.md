---
title: ChartGroup.HasDropLines Property (Excel)
keywords: vbaxl10.chm568079
f1_keywords:
- vbaxl10.chm568079
ms.prod: excel
api_name:
- Excel.ChartGroup.HasDropLines
ms.assetid: cc0d188d-51ba-951d-7063-10820e5e4a42
ms.date: 06/08/2017
---


# ChartGroup.HasDropLines Property (Excel)

 **True** if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDropLines**

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


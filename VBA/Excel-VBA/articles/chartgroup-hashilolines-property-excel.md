---
title: ChartGroup.HasHiLoLines Property (Excel)
keywords: vbaxl10.chm568080
f1_keywords:
- vbaxl10.chm568080
ms.prod: excel
api_name:
- Excel.ChartGroup.HasHiLoLines
ms.assetid: ea743b83-8a3c-7ce1-6659-9a25ebb8eeae
ms.date: 06/08/2017
---


# ChartGroup.HasHiLoLines Property (Excel)

 **True** if the line chart has high-low lines. Applies only to line charts. Read/write **Boolean** .


## Syntax

 _expression_ . **HasHiLoLines**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example turns on high-low lines for chart group one in Chart1 and then sets line style, weight, and color. The example should be run on a 2-D line chart that has three series of stock-quote-like data (high-low-close).


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


---
title: ChartGroup.SplitValue Property (Excel)
keywords: vbaxl10.chm568098
f1_keywords:
- vbaxl10.chm568098
ms.prod: excel
api_name:
- Excel.ChartGroup.SplitValue
ms.assetid: a7cab670-1510-5334-f11b-12dc8cc13570
ms.date: 06/08/2017
---


# ChartGroup.SplitValue Property (Excel)

Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write  **Variant** .


## Syntax

 _expression_ . **SplitValue**

 _expression_ A variable that represents a **ChartGroup** object.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section.


```vb
With Worksheets(1).ChartObjects(1).Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


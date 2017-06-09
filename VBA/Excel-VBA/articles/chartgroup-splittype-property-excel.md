---
title: ChartGroup.SplitType Property (Excel)
keywords: vbaxl10.chm568097
f1_keywords:
- vbaxl10.chm568097
ms.prod: excel
api_name:
- Excel.ChartGroup.SplitType
ms.assetid: c65ca7a4-59b1-6b15-116a-f76007fbd4be
ms.date: 06/08/2017
---


# ChartGroup.SplitType Property (Excel)

Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write  **[XlChartSplitType](xlchartsplittype-enumeration-excel.md)** .


## Syntax

 _expression_ . **SplitType**

 _expression_ A variable that represents a **ChartGroup** object.


## Remarks





| **XlChartSplitType** can be one of these **XlChartSplitType** constants.|
| **xlSplitByCustomSplit**|
| **xlSplitByPercentValue**|
| **xlSplitByPosition**|
| **xlSplitByValue**|

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


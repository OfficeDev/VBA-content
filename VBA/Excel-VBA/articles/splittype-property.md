---
title: SplitType Property
keywords: vbagr10.chm3077588
f1_keywords:
- vbagr10.chm3077588
ms.prod: excel
api_name:
- Excel.SplitType
ms.assetid: e6af8aac-bd1f-9e00-abd7-54e49623d536
ms.date: 06/08/2017
---


# SplitType Property

Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write XlChartSplitType .



|XlChartSplitType can be one of these XlChartSplitType constants.|
| **xlSplitByPercentValue**|
| **xlSplitByValue**|
| **xlSplitByCustomSplit**|
| **xlSplitByPosition**|

 _expression_. **SplitType**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section.


```vb
With myChart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
End With
```



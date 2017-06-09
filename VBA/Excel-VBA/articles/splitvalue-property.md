---
title: SplitValue Property
keywords: vbagr10.chm5208024
f1_keywords:
- vbagr10.chm5208024
ms.prod: excel
api_name:
- Excel.SplitValue
ms.assetid: 3200801a-9464-6bde-59a2-0a8baafcb8ff
ms.date: 06/08/2017
---


# SplitValue Property

Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write  **Variant**.


## Example

This example must be run on either a pie of pie chart or a bar of pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section.


```vb
With myChart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
End With
```



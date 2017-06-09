---
title: ChartGroup.GapWidth Property (Excel)
keywords: vbaxl10.chm568078
f1_keywords:
- vbaxl10.chm568078
ms.prod: excel
api_name:
- Excel.ChartGroup.GapWidth
ms.assetid: 2bf93d07-9181-f43c-5a0f-9350fc1ebd62
ms.date: 06/08/2017
---


# ChartGroup.GapWidth Property (Excel)

Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long** .


## Syntax

 _expression_ . **GapWidth**

 _expression_ A variable that represents a **ChartGroup** object.


## Remarks

The value of this property must be between 0 and 500.


## Example

This example sets the space between column clusters in Chart1 to be 50 percent of the column width.


```vb
Charts("Chart1").ChartGroups(1).GapWidth = 50
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)


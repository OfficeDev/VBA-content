---
title: GapWidth Property
keywords: vbagr10.chm65587
f1_keywords:
- vbagr10.chm65587
ms.prod: excel
api_name:
- Excel.GapWidth
ms.assetid: d00b2a8b-76a0-1dbe-537d-bb55f3a069c9
ms.date: 06/08/2017
---


# GapWidth Property

Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. The value of this property must be between 0 and 500. Read/write  **Long**.

Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the specified chart. The value of this property must be between 5 and 200. Read/write  **Long**.

## Example

This example sets the space between column clusters to be 50 percent of the column width.


```
myChart.ChartGroups(1).GapWidth = 50
```



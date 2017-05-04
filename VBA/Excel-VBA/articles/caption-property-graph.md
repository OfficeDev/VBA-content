---
title: Caption Property (Graph)
keywords: vbagr10.chm3076968
f1_keywords:
- vbagr10.chm3076968
ms.prod: EXCEL
ms.assetid: 37d9afab-873c-c026-fb76-33987aa103b8
---


# Caption Property (Graph)

Returns or sets the title text for the object. Read/write String.

 _expression_. **Caption**

 _expression_ Required. An expression that returns one of the above objects.


## Example

This example adds the title "Annual Salary Figures" to the chart.


```vb
myChart.HasTitle = True 
myChart.ChartTitle.Caption = "Annual Salary Figures" 

```



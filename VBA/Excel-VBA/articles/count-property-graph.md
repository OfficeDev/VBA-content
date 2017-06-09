---
title: Count Property (Graph)
keywords: vbagr10.chm65654
f1_keywords:
- vbagr10.chm65654
ms.prod: excel
ms.assetid: 35eab4b7-6b48-c037-6d25-1d3a0016a528
ms.date: 06/08/2017
---


# Count Property (Graph)

Returns the number of objects in the specified collection. Read-only  **Long**.


## Example

This example displays the number of chart groups in the chart.


```vb
MsgBox "The chart contains " &; _ 
 myChart.ChartGroups.Count &; _ 
 " chart groups."
```



---
title: ClearContents Method
keywords: vbagr10.chm65649
f1_keywords:
- vbagr10.chm65649
ms.prod: excel
api_name:
- Excel.ClearContents
ms.assetid: 8bf70623-e644-e45e-1b1e-565fe6acd223
ms.date: 06/08/2017
---


# ClearContents Method

ClearContents method as it applies to the  **ChartArea** object.

Clears the data from a chart but leaves the formatting.

 _expression_. **ClearContents**

 _expression_ Required. An expression that returns one of the above objects.
ClearContents method as it applies to the  **Range** object.
Clears the formulas from the range.
 _expression_. **ClearContents**
 _expression_ Required. An expression that returns one of the above objects.

## Example

This example clears the formulas from cells A1:G37 on the datasheet but leaves the formatting intact.


```
myChart.Application.DataSheet.Range("A1:G37").ClearContents
```

This example clears the chart data from a chart but leaves the formatting intact.




```
myChart.ChartArea.ClearContents
```



---
title: ChartArea.ClearContents Method (Excel)
keywords: vbaxl10.chm620078
f1_keywords:
- vbaxl10.chm620078
ms.prod: excel
api_name:
- Excel.ChartArea.ClearContents
ms.assetid: 3c3c07a0-9dc1-6019-5262-e1acba7917a1
ms.date: 06/08/2017
---


# ChartArea.ClearContents Method (Excel)

Clears the data from a chart but leaves the formatting.


## Syntax

 _expression_ . **ClearContents**

 _expression_ A variable that represents a **ChartArea** object.


### Return Value

Variant


## Example

This example clears the chart data from Chart1 but leaves the formatting intact.


```vb
Charts("Chart1").ChartArea.ClearContents
```


## See also


#### Concepts


[ChartArea Object](chartarea-object-excel.md)


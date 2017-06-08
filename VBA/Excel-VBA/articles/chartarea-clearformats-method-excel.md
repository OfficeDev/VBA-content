---
title: ChartArea.ClearFormats Method (Excel)
keywords: vbaxl10.chm620082
f1_keywords:
- vbaxl10.chm620082
ms.prod: excel
api_name:
- Excel.ChartArea.ClearFormats
ms.assetid: 0af0bba7-6fb8-d221-7b1f-ba7c40ae1687
ms.date: 06/08/2017
---


# ChartArea.ClearFormats Method (Excel)

Clears the formatting of the object.


## Syntax

 _expression_ . **ClearFormats**

 _expression_ A variable that represents a **ChartArea** object.


### Return Value

Variant


## Example

This example clears the formatting from embedded chart one on Sheet1.


```vb
Worksheets("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats
```


## See also


#### Concepts


[ChartArea Object](chartarea-object-excel.md)


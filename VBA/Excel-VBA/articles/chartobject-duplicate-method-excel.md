---
title: ChartObject.Duplicate Method (Excel)
keywords: vbaxl10.chm494080
f1_keywords:
- vbaxl10.chm494080
ms.prod: excel
api_name:
- Excel.ChartObject.Duplicate
ms.assetid: f43de123-c113-ca5d-6cf7-71f7d08f7e88
ms.date: 06/08/2017
---


# ChartObject.Duplicate Method (Excel)

Duplicates the object and returns a reference to the new copy.


## Syntax

 _expression_ . **Duplicate**

 _expression_ A variable that represents a **ChartObject** object.


### Return Value

Object


## Example

This example duplicates embedded chart one on Sheet1 and then selects the copy.


```vb
Set dChart = Worksheets("Sheet1").ChartObjects(1).Duplicate 
dChart.Select
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)


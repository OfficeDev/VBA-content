---
title: ChartObjects.Duplicate Method (Excel)
keywords: vbaxl10.chm497079
f1_keywords:
- vbaxl10.chm497079
ms.prod: excel
api_name:
- Excel.ChartObjects.Duplicate
ms.assetid: 085e07e1-7b08-befb-1351-b9de3df26ddc
ms.date: 06/08/2017
---


# ChartObjects.Duplicate Method (Excel)

Duplicates the object and returns a reference to the new copy.


## Syntax

 _expression_ . **Duplicate**

 _expression_ A variable that represents a **ChartObjects** object.


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


[ChartObjects Object](chartobjects-object-excel.md)


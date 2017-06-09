---
title: PivotCache.Refresh Method (Excel)
keywords: vbaxl10.chm227080
f1_keywords:
- vbaxl10.chm227080
ms.prod: excel
api_name:
- Excel.PivotCache.Refresh
ms.assetid: 2833d199-342c-9e2e-d1f8-88c33a74bac6
ms.date: 06/08/2017
---


# PivotCache.Refresh Method (Excel)

Causes the specified chart to be redrawn immediately.


## Syntax

 _expression_ . **Refresh**

 _expression_ A variable that represents a **PivotCache** object.


## Example

This example refreshes the PivotTable cache for the first PivotTable report on the first worksheet in a workbook.


```vb
Worksheets(1).PivotTables(1).PivotCache.Refresh
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)


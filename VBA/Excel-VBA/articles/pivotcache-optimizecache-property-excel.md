---
title: PivotCache.OptimizeCache Property (Excel)
keywords: vbaxl10.chm227078
f1_keywords:
- vbaxl10.chm227078
ms.prod: excel
api_name:
- Excel.PivotCache.OptimizeCache
ms.assetid: 4aedf3bb-e15a-439c-5987-ea16cc233a7c
ms.date: 06/08/2017
---


# PivotCache.OptimizeCache Property (Excel)

 **True** if the PivotTable cache is optimized when it's constructed. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **OptimizeCache**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

Cache optimization results in additional queries and degrades initial performance of the PivotTable report.

For OLE DB data sources, this property is read-only and always returns  **False** .


## Example

This example causes the PivotTable cache for the first PivotTable report on worksheet one to be optimized when it's constructed.


```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotCache.OptimizeCache = True
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)


---
title: PivotTable.SubtotalHiddenPageItems Property (Excel)
keywords: vbaxl10.chm235118
f1_keywords:
- vbaxl10.chm235118
ms.prod: excel
api_name:
- Excel.PivotTable.SubtotalHiddenPageItems
ms.assetid: bb3c7e54-1894-a1b6-e2d0-cf6097bd4875
ms.date: 06/08/2017
---


# PivotTable.SubtotalHiddenPageItems Property (Excel)

 **True** if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **SubtotalHiddenPageItems**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

For OLAP data sources, the value is always  **True** .


## Example

This example sets the first PivotTable report on worksheet one to exclude hidden page field items in subtotals.


```vb
Worksheets(1).PivotTables("Pivot1").SubtotalHiddenPageItems = True
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


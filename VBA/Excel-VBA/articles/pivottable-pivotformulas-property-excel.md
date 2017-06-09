---
title: PivotTable.PivotFormulas Property (Excel)
keywords: vbaxl10.chm235116
f1_keywords:
- vbaxl10.chm235116
ms.prod: excel
api_name:
- Excel.PivotTable.PivotFormulas
ms.assetid: fceade1d-7aa1-85c1-ca74-89460ffa6dff
ms.date: 06/08/2017
---


# PivotTable.PivotFormulas Property (Excel)

Returns a  **[PivotFormulas](pivotformulas-object-excel.md)** object that represents the collection of formulas for the specified PivotTable report. Read-only.


## Syntax

 _expression_ . **PivotFormulas**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

For OLAP data sources, this property returns an empty collection.


## Example


```vb
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas 
 r = r + 1 
 Cells(r, 1).Value = pf.Formula 
Next
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


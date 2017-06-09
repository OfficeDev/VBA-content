---
title: PivotTable.CommitChanges Method (Excel)
keywords: vbaxl10.chm235192
f1_keywords:
- vbaxl10.chm235192
ms.prod: excel
api_name:
- Excel.PivotTable.CommitChanges
ms.assetid: f64031c6-8309-7c8a-5786-949d2ec10dea
ms.date: 06/08/2017
---


# PivotTable.CommitChanges Method (Excel)

Performs a commit operation on the data source of a PivotTable report based on an OLAP data source.


## Syntax

 _expression_ . **CommitChanges**

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Return Value

Nothing


## Remarks

The  **CommitChanges** method sends a **COMMIT TRANSACTION** statement to the OLAP server, and clears all cells that were edited by entering a value, but will not clear formulas in value cells. This method generates a run-time error if it is executed on a PivotTable report based on a non-OLAP data source.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


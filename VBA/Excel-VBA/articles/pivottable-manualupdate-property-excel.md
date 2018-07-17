---
title: PivotTable.ManualUpdate Property (Excel)
keywords: vbaxl10.chm235112
f1_keywords:
- vbaxl10.chm235112
ms.prod: excel
api_name:
- Excel.PivotTable.ManualUpdate
ms.assetid: 7686a4d0-720c-949a-d6a1-ba2fdea82340
ms.date: 06/08/2017
---


# PivotTable.ManualUpdate Property (Excel)

 **True** if the PivotTable report is recalculated only at the user's request. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **ManualUpdate**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This property is set to  **False** immediately after your program terminates and after you execute the statement in the Immediate window of the Microsoft Visual Basic Editor.


## Example

This example causes the PivotTable report to be recalculated only at the user's request.


```vb
Worksheets(1).PivotTables("Pivot1").ManualUpdate = True
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


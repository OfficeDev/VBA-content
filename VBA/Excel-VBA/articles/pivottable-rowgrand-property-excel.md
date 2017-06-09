---
title: PivotTable.RowGrand Property (Excel)
keywords: vbaxl10.chm235094
f1_keywords:
- vbaxl10.chm235094
ms.prod: excel
api_name:
- Excel.PivotTable.RowGrand
ms.assetid: 9d016b8d-4c2b-86a3-bcf1-a9a7356b825d
ms.date: 06/08/2017
---


# PivotTable.RowGrand Property (Excel)

 **True** if the PivotTable report shows grand totals for rows. Read/write **Boolean** .


## Syntax

 _expression_ . **RowGrand**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example sets the PivotTable report to show grand totals for rows.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.RowGrand = True
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


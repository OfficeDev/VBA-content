---
title: PivotTable.RowRange Property (Excel)
keywords: vbaxl10.chm235095
f1_keywords:
- vbaxl10.chm235095
ms.prod: excel
api_name:
- Excel.PivotTable.RowRange
ms.assetid: 3b586599-9b2a-d0fc-c205-b8e3c6e7074f
ms.date: 06/08/2017
---


# PivotTable.RowRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range including the row area on the PivotTable report. Read-only.


## Syntax

 _expression_ . **RowRange**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example selects the row headers on the PivotTable report.


```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.RowRange.Select
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


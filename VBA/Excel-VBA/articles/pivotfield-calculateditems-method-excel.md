---
title: PivotField.CalculatedItems Method (Excel)
keywords: vbaxl10.chm240100
f1_keywords:
- vbaxl10.chm240100
ms.prod: excel
api_name:
- Excel.PivotField.CalculatedItems
ms.assetid: 89818448-9a1e-0dcd-5e0f-479bf051d590
ms.date: 06/08/2017
---


# PivotField.CalculatedItems Method (Excel)

Returns a  **[CalculatedItems](calculateditems-object-excel.md)** collection that represents all the calculated items in the specified PivotTable report. Read-only.


## Syntax

 _expression_ . **CalculatedItems**

 _expression_ A variable that represents a **PivotField** object.


### Return Value

CalculatedItems


## Remarks

For OLAP data sources, this method returns a zero-length collection.


## Example

This example creates a list of calculated items and their formulas.


```vb
Set pt = Worksheets(1).PivotTables(1) 
For Each ci In pt.PivotFields("Sales").CalculatedItems 
 r = r + 1 
 With Worksheets(2) 
 .Cells(r, 1).Value = ci.Name 
 .Cells(r, 2).Value = ci.Formula 
 End With 
Next
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)


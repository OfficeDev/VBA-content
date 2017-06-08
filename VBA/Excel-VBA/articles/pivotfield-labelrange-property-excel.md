---
title: PivotField.LabelRange Property (Excel)
keywords: vbaxl10.chm240084
f1_keywords:
- vbaxl10.chm240084
ms.prod: excel
api_name:
- Excel.PivotField.LabelRange
ms.assetid: be06bf39-d970-316e-6833-65efde85ddc8
ms.date: 06/08/2017
---


# PivotField.LabelRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the cell (or cells) that contain the field label. Read-only


## Syntax

 _expression_ . **LabelRange**

 _expression_ A variable that represents a **PivotField** object.


## Example

This example selects the field button for the field named "ORDER_DATE."


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set pvtField = pvtTable.PivotFields("ORDER_DATE") 
Worksheets("Sheet1").Activate 
pvtField.LabelRange.Select
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)


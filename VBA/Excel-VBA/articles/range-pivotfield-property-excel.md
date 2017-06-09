---
title: Range.PivotField Property (Excel)
keywords: vbaxl10.chm144175
f1_keywords:
- vbaxl10.chm144175
ms.prod: excel
api_name:
- Excel.Range.PivotField
ms.assetid: 56003d2d-60cd-abd2-455e-4a4d3616a615
ms.date: 06/08/2017
---


# Range.PivotField Property (Excel)

Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the PivotTable field containing the upper-left corner of the specified range.


## Syntax

 _expression_ . **PivotField**

 _expression_ A variable that represents a **Range** object.


## Example

This example displays the name of the PivotTable field that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the field " &; _ 
 ActiveCell.PivotField.Name
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


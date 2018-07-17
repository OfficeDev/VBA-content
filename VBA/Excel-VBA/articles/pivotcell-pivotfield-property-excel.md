---
title: PivotCell.PivotField Property (Excel)
keywords: vbaxl10.chm692076
f1_keywords:
- vbaxl10.chm692076
ms.prod: excel
api_name:
- Excel.PivotCell.PivotField
ms.assetid: a1217848-e3b0-0e92-168b-3a9c21245380
ms.date: 06/08/2017
---


# PivotCell.PivotField Property (Excel)

Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the PivotTable field containing the upper-left corner of the specified range.


## Syntax

 _expression_ . **PivotField**

 _expression_ A variable that represents a **PivotCell** object.


## Example

This example displays the name of the PivotTable field that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the field " &; _ 
 ActiveCell.PivotField.Name
```


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)


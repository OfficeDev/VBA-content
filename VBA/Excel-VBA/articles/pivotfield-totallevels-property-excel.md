---
title: PivotField.TotalLevels Property (Excel)
keywords: vbaxl10.chm240097
f1_keywords:
- vbaxl10.chm240097
ms.prod: excel
api_name:
- Excel.PivotField.TotalLevels
ms.assetid: fa50c186-5f6d-41f4-6382-37135159347c
ms.date: 06/08/2017
---


# PivotField.TotalLevels Property (Excel)

Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based,  **TotalLevels** returns the value 1. Read-only **Long** .


## Syntax

 _expression_ . **TotalLevels**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

All fields in a set of grouped fields have the same  **TotalLevels** value.


## Example

This example displays the total number of fields in the group that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
MsgBox "This group has " &; _ 
 ActiveCell.PivotField.TotalLevels &; " levels
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)


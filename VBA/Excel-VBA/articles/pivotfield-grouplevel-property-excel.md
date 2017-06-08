---
title: PivotField.GroupLevel Property (Excel)
keywords: vbaxl10.chm240082
f1_keywords:
- vbaxl10.chm240082
ms.prod: excel
api_name:
- Excel.PivotField.GroupLevel
ms.assetid: fc017652-bded-4655-03df-79cfa733b12e
ms.date: 06/08/2017
---


# PivotField.GroupLevel Property (Excel)

Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.


## Syntax

 _expression_ . **GroupLevel**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This property is not available for OLAP data sources.

The highest-level parent field (leftmost parent field) is level one, its child is level two, and so on.


## Example

This example displays a message box if the field that contains the active cell is the highest-level parent field.


```vb
Worksheets("Sheet1").Activate 
If ActiveCell.PivotField.GroupLevel = 1 Then 
 MsgBox "This is the highest-level parent field." 
End If
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)


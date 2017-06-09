---
title: PivotTable.SelectionMode Property (Excel)
keywords: vbaxl10.chm235125
f1_keywords:
- vbaxl10.chm235125
ms.prod: excel
api_name:
- Excel.PivotTable.SelectionMode
ms.assetid: 692c31b9-01a4-2a49-65c9-66c14ab6aa7c
ms.date: 06/08/2017
---


# PivotTable.SelectionMode Property (Excel)

Returns or sets the PivotTable report structured selection mode. Read/write  **[XlPTSelectionMode](xlptselectionmode-enumeration-excel.md)** .


## Syntax

 _expression_ . **SelectionMode**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks



| **XlPTSelectionMode** can be one of these **XlPTSelectionMode** constants.|
| **xlBlanks**|
| **xlButton**|
| **xlDataAndLabel**|
| **xlDataOnly**|
| **xlFirstRow**|
| **xlLabelOnly**|
| **xlOrigin**|
If the PivotTable field isn't in outline form, specifying the sum of any of the constants and  **xlFirstRow** is equivalent to specifying the constant alone.


## Example

This example enables structured selection mode and then sets the first PivotTable report on worksheet one to allow only data to be selected.


```vb
Application.PivotTableSelection = True 
Worksheets(1).PivotTables(1).SelectionMode = xlDataOnly
```

In this example, the PivotTable report is in outline form.




```vb
Application.PivotTableSelection = True 
Worksheets(1).PivotTables(1).SelectionMode = _ 
 xlDataOnly + xlFirstRow
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


---
title: PivotTable.TableRange1 Property (Excel)
keywords: vbaxl10.chm235098
f1_keywords:
- vbaxl10.chm235098
ms.prod: excel
api_name:
- Excel.PivotTable.TableRange1
ms.assetid: 4dfea643-3299-82ee-a770-b961904eec7f
ms.date: 06/08/2017
---


# PivotTable.TableRange1 Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.


## Syntax

 _expression_ . **TableRange1**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The  **[TableRange2](pivottable-tablerange2-property-excel.md)** property includes page fields.


## Example

This example selects all of the PivotTable report except its page fields.


```vb
Worksheets("Sheet1").Activate 
Range("A3").PivotTable.TableRange1.Select
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


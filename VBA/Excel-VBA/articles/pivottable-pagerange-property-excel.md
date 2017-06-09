---
title: PivotTable.PageRange Property (Excel)
keywords: vbaxl10.chm235087
f1_keywords:
- vbaxl10.chm235087
ms.prod: excel
api_name:
- Excel.PivotTable.PageRange
ms.assetid: 05629703-c43f-282c-e4da-22c95094e15b
ms.date: 06/08/2017
---


# PivotTable.PageRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range that contains the page area in the PivotTable report. Read-only.


## Syntax

 _expression_ . **PageRange**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example selects the page headers in the PivotTable report.


```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.PageRange.Select
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


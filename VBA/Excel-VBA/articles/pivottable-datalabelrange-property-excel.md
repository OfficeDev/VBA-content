---
title: PivotTable.DataLabelRange Property (Excel)
keywords: vbaxl10.chm235080
f1_keywords:
- vbaxl10.chm235080
ms.prod: excel
api_name:
- Excel.PivotTable.DataLabelRange
ms.assetid: 9a4a6ee0-f918-2dd3-f423-e5ced6fdba20
ms.date: 06/08/2017
---


# PivotTable.DataLabelRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range that contains the labels for the data fields in the PivotTable report. Read-only.


## Syntax

 _expression_ . **DataLabelRange**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example selects the data field labels in the PivotTable report.


```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.DataLabelRange.Select
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)


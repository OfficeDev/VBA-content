---
title: Range.Table Method (Excel)
keywords: vbaxl10.chm144208
f1_keywords:
- vbaxl10.chm144208
ms.prod: excel
api_name:
- Excel.Range.Table
ms.assetid: 804b0e1d-e92d-387d-1054-90643bfd16ff
ms.date: 06/08/2017
---


# Range.Table Method (Excel)

Creates a data table based on input values and formulas that you define on a worksheet.


## Syntax

 _expression_ . **Table**( **_RowInput_** , **_ColumnInput_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowInput_|Optional| **Variant**|A single cell to use as the row input for your table.|
| _ColumnInput_|Optional| **Variant**|A single cell to use as the column input for your table.|

### Return Value

Variant


## Remarks

Use data tables to perform a what-if analysis by changing certain constant values on your worksheet to see how values in other cells are affected.


## Example

This example creates a formatted multiplication table in cells A1:K11 on Sheet1.


```vb
Set dataTableRange = Worksheets("Sheet1").Range("A1:K11") 
Set rowInputCell = Worksheets("Sheet1").Range("A12") 
Set columnInputCell = Worksheets("Sheet1").Range("A13") 
 
Worksheets("Sheet1").Range("A1").Formula = "=A12*A13" 
For i = 2 To 11 
 Worksheets("Sheet1").Cells(i, 1) = i - 1 
 Worksheets("Sheet1").Cells(1, i) = i - 1 
Next i 
dataTableRange.Table rowInputCell, columnInputCell 
With Worksheets("Sheet1").Range("A1").CurrentRegion 
 .Rows(1).Font.Bold = True 
 .Columns(1).Font.Bold = True 
 .Columns.AutoFit 
End With
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


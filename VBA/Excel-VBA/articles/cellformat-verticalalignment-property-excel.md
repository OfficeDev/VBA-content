---
title: CellFormat.VerticalAlignment Property (Excel)
keywords: vbaxl10.chm676081
f1_keywords:
- vbaxl10.chm676081
ms.prod: excel
api_name:
- Excel.CellFormat.VerticalAlignment
ms.assetid: c901dff3-3f0a-1f54-250e-c03b9e32c819
ms.date: 06/08/2017
---


# CellFormat.VerticalAlignment Property (Excel)

Returns or sets a  **Variant** value that represents the vertical alignment of the specified object.


## Syntax

 _expression_ . **VerticalAlignment**

 _expression_ A variable that represents a **CellFormat** object.


## Remarks

The value of this property can be set to one of the following constants:



| **xlBottom**|
| **xlCenter**|
| **xlDistributed**|
| **xlJustify**|
| **xlTop**|

## Example

This example sets the height of row 2 on Sheet1 to twice the standard height and then centers the contents of the row vertically.


```vb
Worksheets("Sheet1").Rows(2).RowHeight = _ 
 2 * Worksheets("Sheet1").StandardHeight 
Worksheets("Sheet1").Rows(2).VerticalAlignment = xlVAlignCenter 

```


## See also


#### Concepts


[CellFormat Object](cellformat-object-excel.md)


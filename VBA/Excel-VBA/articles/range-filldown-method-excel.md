---
title: Range.FillDown Method (Excel)
keywords: vbaxl10.chm144124
f1_keywords:
- vbaxl10.chm144124
ms.prod: excel
api_name:
- Excel.Range.FillDown
ms.assetid: bb7c0b2d-8dd9-13e5-b90a-b2708935afa9
ms.date: 06/08/2017
---


# Range.FillDown Method (Excel)

Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.


## Syntax

 _expression_ . **FillDown**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example fills the range A1:A10 on Sheet1, based on the contents of cell A1.


```vb
Worksheets("Sheet1").Range("A1:A10").FillDown
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


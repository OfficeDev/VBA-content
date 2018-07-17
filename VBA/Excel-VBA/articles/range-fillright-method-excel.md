---
title: Range.FillRight Method (Excel)
keywords: vbaxl10.chm144126
f1_keywords:
- vbaxl10.chm144126
ms.prod: excel
api_name:
- Excel.Range.FillRight
ms.assetid: b0b9a3a5-5f8c-327e-fb41-dec5c1a2f2b3
ms.date: 06/08/2017
---


# Range.FillRight Method (Excel)

Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.


## Syntax

 _expression_ . **FillRight**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example fills the range A1:M1 on Sheet1, based on the contents of cell A1.


```vb
Worksheets("Sheet1").Range("A1:M1").FillRight
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


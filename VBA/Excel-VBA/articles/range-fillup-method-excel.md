---
title: Range.FillUp Method (Excel)
keywords: vbaxl10.chm144127
f1_keywords:
- vbaxl10.chm144127
ms.prod: excel
api_name:
- Excel.Range.FillUp
ms.assetid: 52498f52-95f9-5993-7c44-76cd8b696074
ms.date: 06/08/2017
---


# Range.FillUp Method (Excel)

Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.


## Syntax

 _expression_ . **FillUp**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example fills the range A1:A10 on Sheet1, based on the contents of cell A10.


```vb
Worksheets("Sheet1").Range("A1:A10").FillUp
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


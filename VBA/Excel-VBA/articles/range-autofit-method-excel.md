---
title: Range.AutoFit Method (Excel)
keywords: vbaxl10.chm144085
f1_keywords:
- vbaxl10.chm144085
ms.prod: excel
api_name:
- Excel.Range.AutoFit
ms.assetid: 53a35cd3-00e7-f9f5-2cd2-8492d7814a11
ms.date: 06/08/2017
---


# Range.AutoFit Method (Excel)

Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.


## Syntax

 _expression_ . **AutoFit**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Remarks

The  **Range** object must be a row or a range of rows, or a column or a range of columns. Otherwise, this method generates an error.

One unit of column width is equal to the width of one character in the Normal style.


## Example

This example changes the width of columns A through I on Sheet1 to achieve the best fit.


```vb
Worksheets("Sheet1").Columns("A:I").AutoFit
```

This example changes the width of columns A through E on Sheet1 to achieve the best fit, based only on the contents of cells A1:E1.




```vb
Worksheets("Sheet1").Range("A1:E1").Columns.AutoFit
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


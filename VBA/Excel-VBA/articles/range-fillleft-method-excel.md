---
title: Range.FillLeft Method (Excel)
keywords: vbaxl10.chm144125
f1_keywords:
- vbaxl10.chm144125
ms.prod: excel
api_name:
- Excel.Range.FillLeft
ms.assetid: 42722b18-8b40-c27b-8bca-ef180cf0f636
ms.date: 06/08/2017
---


# Range.FillLeft Method (Excel)

Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.


## Syntax

 _expression_ . **FillLeft**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example fills the range A1:M1 on Sheet1, based on the contents of cell M1.


```vb
Worksheets("Sheet1").Range("A1:M1").FillLeft
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


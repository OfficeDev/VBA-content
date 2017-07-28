---
title: Range.ClearContents Method (Excel)
keywords: vbaxl10.chm144095
f1_keywords:
- vbaxl10.chm144095
ms.prod: excel
api_name:
- Excel.Range.ClearContents
ms.assetid: 8c957fdd-e99d-ca0e-7d2c-4cb1db62639a
ms.date: 06/08/2017
---


# Range.ClearContents Method (Excel)

Clears formulas and values from the range.


## Syntax

 _expression_ . **ClearContents**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example clears formulas and values from cells A1:G37 on `Sheet1`, but leaves the cell formatting and conditional formatting intact.


```vb
Worksheets("Sheet1").Range("A1:G37").ClearContents
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


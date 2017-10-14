---
title: Range.FormulaArray Property (Excel)
keywords: vbaxl10.chm144133
f1_keywords:
- vbaxl10.chm144133
ms.prod: excel
api_name:
- Excel.Range.FormulaArray
ms.assetid: a0c8bafb-294c-32ff-0591-1a798aebb4b4
ms.date: 06/08/2017
---


# Range.FormulaArray Property (Excel)

Returns or sets the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns  **null** . Read/write **Variant** .


## Syntax

 _expression_ . **FormulaArray**

 _expression_ . A variable that represents a **Range** object.


## Remarks

The  **FormulaArray** property also has a character limit of 255.


## Example

This example enters the number 3 as an array constant in cells A1:C5 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:C5").FormulaArray = "=3"
```

This example enters the array formula =SUM(A1:C3) in cells E1:E3 on Sheet1.




```vb
Worksheets("Sheet1").Range("E1:E3").FormulaArray = _ 
 "=Sum(A1:C3)"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


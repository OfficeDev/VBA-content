---
title: Range.FormulaR1C1 Property (Excel)
keywords: vbaxl10.chm144137
f1_keywords:
- vbaxl10.chm144137
ms.prod: excel
api_name:
- Excel.Range.FormulaR1C1
ms.assetid: 76f41bf6-94e2-2e6a-30e4-012a735a3374
ms.date: 06/08/2017
---


# Range.FormulaR1C1 Property (Excel)

Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **Variant** .


## Syntax

 _expression_ . **FormulaR1C1**

 _expression_ A variable that represents a **Range** object.


## Remarks

If the cell contains a constant, this property returns the constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## Example

This example sets the formula for cell B1 on Sheet1.


```vb
Worksheets("Sheet1").Range("B1").FormulaR1C1 = "=SQRT(R1C1)"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


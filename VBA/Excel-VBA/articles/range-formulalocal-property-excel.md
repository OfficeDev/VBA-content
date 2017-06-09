---
title: Range.FormulaLocal Property (Excel)
keywords: vbaxl10.chm144136
f1_keywords:
- vbaxl10.chm144136
ms.prod: excel
api_name:
- Excel.Range.FormulaLocal
ms.assetid: c69325d9-d35d-c15a-ae49-7bde2b628428
ms.date: 06/08/2017
---


# Range.FormulaLocal Property (Excel)

Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **Variant** .


## Syntax

 _expression_ . **FormulaLocal**

 _expression_ A variable that represents a **Range** object.


## Remarks

If the cell contains a constant, this property returns that constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## Example

Assume that you enter the formula =SUM(A1:A10) in cell A11 on worksheet one, using the American English version of Microsoft Excel. If you then open the workbook on a computer that's running the German version and run the following example, the example displays the formula =SUMME(A1:A10) in a message box.


```vb
MsgBox Worksheets(1).Range("A11").FormulaLocal
```


## See also


#### Concepts


[Range Object](range-object-excel.md)


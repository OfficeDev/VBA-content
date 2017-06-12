---
title: ChartTitle.FormulaR1C1 Property (Word)
keywords: vbawd10.chm65273896
f1_keywords:
- vbawd10.chm65273896
ms.prod: word
api_name:
- Word.ChartTitle.FormulaR1C1
ms.assetid: 00df6397-0c7c-4b44-4e18-780656c4a60a
ms.date: 06/08/2017
---


# ChartTitle.FormulaR1C1 Property (Word)

Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **String** .


## Syntax

 _expression_ . **FormulaR1C1**

 _expression_ A variable that represents a **[ChartTitle](charttitle-object-word.md)** object.


## Remarks

If the cell contains a constant, this property returns the constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


[ChartTitle Object](charttitle-object-word.md)


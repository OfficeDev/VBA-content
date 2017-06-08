---
title: AxisTitle.FormulaLocal Property (Word)
keywords: vbawd10.chm98238506
f1_keywords:
- vbawd10.chm98238506
ms.prod: word
api_name:
- Word.AxisTitle.FormulaLocal
ms.assetid: 05995afc-63a6-58b8-e5e1-380476dfe8ac
ms.date: 06/08/2017
---


# AxisTitle.FormulaLocal Property (Word)

Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String** .


## Syntax

 _expression_ . **FormulaLocal**

 _expression_ A variable that represents an **[AxisTitle](axistitle-object-word.md)** object.


## Remarks

If the cell contains a constant, this property returns that constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign (=)).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-word.md)


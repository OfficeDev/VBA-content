---
title: Series.FormulaLocal Property (Word)
keywords: vbawd10.chm123732231
f1_keywords:
- vbawd10.chm123732231
ms.prod: word
api_name:
- Word.Series.FormulaLocal
ms.assetid: 45d948ca-8519-1941-9dea-6a5e9fe42819
ms.date: 06/08/2017
---


# Series.FormulaLocal Property (Word)

Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String** .


## Syntax

 _expression_ . **FormulaLocal**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

If the cell contains a constant, this property returns that constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Word verifies whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


[Series Object](series-object-word.md)


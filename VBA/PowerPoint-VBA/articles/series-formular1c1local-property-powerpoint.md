---
title: Series.FormulaR1C1Local Property (PowerPoint)
keywords: vbapp10.chm65801
f1_keywords:
- vbapp10.chm65801
ms.prod: powerpoint
api_name:
- PowerPoint.Series.FormulaR1C1Local
ms.assetid: cb00cca5-b540-6083-7fc5-2d2d6a58719f
ms.date: 06/08/2017
---


# Series.FormulaR1C1Local Property (PowerPoint)

Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write  **String**.


## Syntax

 _expression_. **FormulaR1C1Local**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Remarks

If the cell contains a constant, this property returns that constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Word verifies whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)


---
title: DisplayUnitLabel.FormulaR1C1 Property (PowerPoint)
keywords: vbapp10.chm699025
f1_keywords:
- vbapp10.chm699025
ms.prod: powerpoint
api_name:
- PowerPoint.DisplayUnitLabel.FormulaR1C1
ms.assetid: e0f17420-080f-6632-3898-df0c5c7693f8
ms.date: 06/08/2017
---


# DisplayUnitLabel.FormulaR1C1 Property (PowerPoint)

Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **String**.


## Syntax

 _expression_. **FormulaR1C1**

 _expression_ A variable that represents a **DisplayUnitLabel** object.


## Remarks

If the cell contains a constant, this property returns the constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft PowerPoint verifies whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one-dimensional or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


[DisplayUnitLabel Object](displayunitlabel-object-powerpoint.md)


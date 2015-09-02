
# PivotFormula.Formula Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.

## Syntax

 _expression_. **Formula**

 _expression_A variable that represents a  **PivotFormula** object.


## Remarks

This property is not available for OLAP data sources.

If the cell contains a constant, this property returns the constant. If the cell is empty, this property returns an empty string. If the cell contains a formula, the  **Formula** property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, Microsoft Excel changes the number format to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula for a multiple-cell range fills all cells in the range with the formula.


## See also


#### Concepts


 [PivotFormula Object](2955dad6-d686-1a83-ab56-76a00272c7e2.md)
#### Other resources


 [PivotFormula Object Members](8cd6a78c-0b06-353c-8a74-83996eed80b2.md)

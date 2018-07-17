---
title: Range.Formula Property (Excel)
keywords: vbaxl10.chm144132
f1_keywords:
- vbaxl10.chm144132
ms.prod: excel
api_name:
- Excel.Range.Formula
ms.assetid: c5be8952-fc3f-bdb3-d4a6-abf9d94eab1e
ms.date: 06/08/2017
---


# Range.Formula Property (Excel)

Returns or sets a  **Variant** value that represents the object's formula in A1-style notation and in the macro language.


## Syntax

 _expression_ . **Formula**

 _expression_ A variable that represents a **Range** object.


## Remarks

This property is not available for OLAP data sources.

If the cell contains a constant, this property returns the constant. If the cell is empty, this property returns an empty string. If the cell contains a formula, the  **Formula** property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign (=)).

If you set the value or formula of a cell to a date, Microsoft Excel verifies that cell is already formatted with one of the date or time number formats. If not, Microsoft Excel changes the number format to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula for a multiple-cell range fills all cells in the range with the formula.


## Example

The following code example sets the formula for cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)

The following code example sets the formula for cell A1 on Sheet1 to display today's date.




```vb
Sub InsertTodaysDate() 
    ' This macro will put today's date in cell A1 on Sheet1 
    Sheets("Sheet1").Select 
    Range("A1").Select 
    Selection.Formula = "=text(now(),""mmm dd yyyy"")" 
    Selection.Columns.AutoFit 
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Range Object](range-object-excel.md)


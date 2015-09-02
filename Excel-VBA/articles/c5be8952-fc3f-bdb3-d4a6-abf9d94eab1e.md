
# Range.Formula Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)
 [About the Contributor](#AboutContributor)


Returns or sets a  **Variant** value that represents the object's formula in A1-style notation and in the macro language.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Formula**

 _expression_A variable that represents a  **Range** object.


## Remarks
<a name="sectionSection1"> </a>

This property is not available for OLAP data sources.

If the cell contains a constant, this property returns the constant. If the cell is empty, this property returns an empty string. If the cell contains a formula, the  **Formula** property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign (=)).

If you set the value or formula of a cell to a date, Microsoft Excel verifies that cell is already formatted with one of the date or time number formats. If not, Microsoft Excel changes the number format to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula for a multiple-cell range fills all cells in the range with the formula.


## Example
<a name="sectionSection2"> </a>

The following code example sets the formula for cell A1 on Sheet1.


```
Worksheets("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
```



 **Sample code provided by:** Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)

The following code example sets the formula for cell A1 on Sheet1 to display today's date.




```
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


 [Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


 [Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)

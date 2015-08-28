
# PageSetup.PrintTitleColumns Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro. Read/write  **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **PrintTitleColumns**

 _expression_A variable that represents a  **PageSetup** object.


## Remarks
<a name="sectionSection1"> </a>

If you specify only part of a column or columns, Microsoft Excel expands the range to full columns.

Set this property to  **False** or to the empty string ("") to turn off title columns.

This property applies only to worksheet pages.


## Example
<a name="sectionSection2"> </a>

This example defines row three as the title row, and it defines columns one through three as the title columns.


```
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows(3).Address 
ActiveSheet.PageSetup.PrintTitleColumns = _ 
 ActiveSheet.Columns("A:C").Address
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PageSetup Object](2fd22df9-5987-f723-04a9-9a3f2e84ac81.md)
#### Other resources


 [PageSetup Object Members](feabe079-cb03-f560-6032-88f5585ec8a8.md)

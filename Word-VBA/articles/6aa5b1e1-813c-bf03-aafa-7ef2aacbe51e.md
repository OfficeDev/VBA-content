
# LineFormat.Pattern Property (Word)

 **Last modified:** July 28, 2015

Returns or sets a value that represents the pattern applied to the specified line. Read/write  **MsoPatternType**.

## Syntax

 _expression_. **Pattern**

 _expression_Required. A variable that represents a  ** [LineFormat](28fabccb-d03f-3466-9d07-ea3ebc4cdd11.md)** object.


## Example

This example adds a patterned line to  _myDocument_.


```
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddLine(10, 100, 250, 0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With
```


## See also


#### Concepts


 [LineFormat Object](28fabccb-d03f-3466-9d07-ea3ebc4cdd11.md)
#### Other resources


 [LineFormat Object Members](775fcd1f-f4be-f607-c63b-4ae952b7c524.md)

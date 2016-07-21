
# Range.IndentLevel Property (Excel)

Returns or sets a  **Variant** value that represents the indent level for the cell or range. Can be an integer from 0 to 15.


## Syntax

 _expression_ . **IndentLevel**

 _expression_ A variable that represents a **Range** object.


## Remarks

Using this property to set the indent level to a number less than 0 (zero) or greater than 15 causes an error.


## Example

This example increases the indent level to 15 in cell A10.


```vb
With Range("A10") 
 .IndentLevel = 15 
End With
```


## See also


#### Concepts


[Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


[Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)


# Cell.ColumnIndex Property (Word)

 **Last modified:** July 28, 2015

Returns the number of the table column that contains the specified cell. Read-only  **Long**.

## Syntax

 _expression_. **ColumnIndex**

 _expression_A variable that represents a  ** [Cell](cbe6ae71-b2da-63a9-1446-0a2f81ab8b14.md)** object.


## Example

This example creates a table in a new document, selects each cell in the first row, and returns the column number that contains the selected cell.


```
Dim docNew As Document 
Dim tableNew As Table 
Dim cellLoop As Cell 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 
For Each cellLoop In tableNew.Rows(1).Cells 
 cellLoop.Select 
 MsgBox "This is column " &amp; cellLoop.ColumnIndex 
Next cellLoop
```

This example returns the column number of the cell that contains the insertion point.




```
Msgbox Selection.Cells(1).ColumnIndex
```


## See also


#### Concepts


 [Cell Object](cbe6ae71-b2da-63a9-1446-0a2f81ab8b14.md)
#### Other resources


 [Cell Object Members](f718bcaa-af8a-682b-f403-6db1aeb9bb73.md)


# DataTable.HasBorderOutline Property (Excel)

 **Last modified:** July 28, 2015

 **True** if the chart data table has outline borders. Read/write **Boolean**.

## Syntax

 _expression_. **HasBorderOutline**

 _expression_A variable that represents a  **DataTable** object.


## Example

This example causes the embedded chart data table to be displayed with an outline border and no cell borders.


```
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
End With
```


## See also


#### Concepts


 [DataTable Object](aca0850b-2e72-cde9-b751-633876e1df99.md)
#### Other resources


 [DataTable Object Members](5a46944b-e7e6-ac7c-6b95-736975a0a3eb.md)

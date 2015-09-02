
# PivotCell.PivotTable Property (Excel)

 **Last modified:** July 28, 2015

Returns a  ** [PivotTable](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)** object that represents the PivotTable report associated with the PivotCell.

## Syntax

 _expression_. **PivotTable**

 _expression_A variable that represents a  **PivotCell** object.


## Example

This example sets the current page for the PivotTable report on Sheet1 to the page named "Canada."


```
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.PivotFields("Country").CurrentPage = "Canada"
```

This example determines the PivotTable report associated with the Sales chart on the active worksheet, and then it sets the page named "Oregon" as the current page for the PivotTable report.




```
Set objPT = _ 
 ActiveSheet.Charts("Sales").PivotLayout.PivotTable 
objPT.PivotFields("State").CurrentPageName = "Oregon"
```


## See also


#### Concepts


 [PivotCell Object](76b8a2dc-90ee-7475-d327-d27cb1e92703.md)
#### Other resources


 [PivotCell Object Members](e486cd5d-3f31-29d4-b811-24fc0aed6803.md)

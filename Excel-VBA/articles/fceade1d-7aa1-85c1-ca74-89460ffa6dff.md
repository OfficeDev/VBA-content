
# PivotTable.PivotFormulas Property (Excel)

Returns a  **[PivotFormulas](7139a4bd-f103-7190-004f-7f2261a4391f.md)** object that represents the collection of formulas for the specified PivotTable report. Read-only.


## Syntax

 _expression_ . **PivotFormulas**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

For OLAP data sources, this property returns an empty collection.


## Example


```vb
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas 
 r = r + 1 
 Cells(r, 1).Value = pf.Formula 
Next
```


## See also


#### Concepts


[PivotTable Object](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)
#### Other resources


[PivotTable Object Members](8e8d1692-cf32-63c6-a1f6-54ddcc2a4964.md)

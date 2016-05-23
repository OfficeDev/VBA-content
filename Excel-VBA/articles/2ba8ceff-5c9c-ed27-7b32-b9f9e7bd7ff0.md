
# PivotItem.RecordCount Property (Excel)

Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .


## Syntax

 _expression_ . **RecordCount**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

This property reflects the transient state of the cache at the time that it's queried. The cache can change between queries.


## Example

This example displays the number of cache records that contain "Kiwi" in the "Products" field.


```vb
MsgBox Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Product").PivotItems("Kiwi").RecordCount
```


## See also


#### Concepts


[PivotItem Object](5829a1d9-0924-9ce8-1120-229e4595285a.md)
#### Other resources


[PivotItem Object Members](dde86683-8c89-2484-cdd0-8c3db0c06f45.md)

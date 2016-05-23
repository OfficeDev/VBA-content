
# PivotCache.RecordCount Property (Excel)

Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .


## Syntax

 _expression_ . **RecordCount**

 _expression_ A variable that represents a **PivotCache** object.


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


[PivotCache Object](c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439.md)
#### Other resources


[PivotCache Object Members](113f1109-e1c9-2c6e-0581-9fba82f278dc.md)

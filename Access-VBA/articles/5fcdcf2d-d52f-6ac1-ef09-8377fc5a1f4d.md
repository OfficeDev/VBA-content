
# PivotCache.RecordCount Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RecordCount**

 _expression_A variable that represents a  **PivotCache** object.


## Remarks
<a name="sectionSection1"> </a>

This property reflects the transient state of the cache at the time that it's queried. The cache can change between queries.


## Example
<a name="sectionSection2"> </a>

This example displays the number of cache records that contain "Kiwi" in the "Products" field.


```
MsgBox Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Product").PivotItems("Kiwi").RecordCount
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PivotCache Object](c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439.md)
#### Other resources


 [PivotCache Object Members](113f1109-e1c9-2c6e-0581-9fba82f278dc.md)

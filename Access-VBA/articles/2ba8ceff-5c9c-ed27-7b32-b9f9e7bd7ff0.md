
# PivotItem.RecordCount Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RecordCount**

 _expression_A variable that represents a  **PivotItem** object.


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


 [PivotItem Object](5829a1d9-0924-9ce8-1120-229e4595285a.md)
#### Other resources


 [PivotItem Object Members](dde86683-8c89-2484-cdd0-8c3db0c06f45.md)


# PivotField.VisibleItemsList Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  **Variant** specifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **VisibleItemsList**

 _expression_A variable that represents a  **PivotField** object.


## Remarks
<a name="sectionSection1"> </a>

This property is applicable to OLAP PivotTables only.


## Example
<a name="sectionSection2"> </a>

This example shows manual, inclusive filtering in an OLAP PivotTable.


```
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] &amp; _ 
.[Country]").VisibleItemsList = Array("[Customer].[Customer Geography].[Country].&amp;[Australia]") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] &amp; _ 
.[State-Province]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] &amp; _ 
.[City]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] &amp; _ 
.[Postal Code]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] &amp; _ 
.[Full Name]").VisibleItemsList = Array("") 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PivotField Object](52784960-e2da-b43a-1e37-2d4dae61c6d8.md)
#### Other resources


 [PivotField Object Members](4a6ea12a-072c-a386-c855-7bf5f6eadd46.md)


# QueryTable.FetchedRowOverflow Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the number of rows returned by the last use of the ** [Refresh](445d74fb-1a9c-bba4-2d53-0ab0caa876da.md)** method is greater than the number of rows available on the worksheet. Read-only **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FetchedRowOverflow**

 _expression_A variable that represents a  **QueryTable** object.


## Remarks
<a name="sectionSection1"> </a>

If you import data using the user interface, data from a Web query or a text query is imported as a  ** [QueryTable](505b84ea-64b3-b4fe-741a-de6884eb69eb.md)** object, while all other external data is imported as a ** [ListObject](46de6c4f-8ce0-0c7d-da59-6e52f5eab612.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a  **ListObject** or a **QueryTable**.

You can use the  ** [QueryTable](fe019d61-654a-9c87-0bf4-30590a1274ca.md)** property of the **ListObject** to access the **FetchedRowOverflow** property.


## Example
<a name="sectionSection2"> </a>

This example refreshes query table one. If the number of rows returned by the query exceeds the number of rows available on the worksheet, an error message is displayed.


```
With Worksheets(1).QueryTables(1) 
 .Refresh 
 If .FetchedRowOverflow Then 
 MsgBox "Query too large: please redefine." 
 End If 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [QueryTable Object](505b84ea-64b3-b4fe-741a-de6884eb69eb.md)
#### Other resources


 [QueryTable Object Members](9a61f024-c1dc-c11b-942f-ff2a6617bdc4.md)

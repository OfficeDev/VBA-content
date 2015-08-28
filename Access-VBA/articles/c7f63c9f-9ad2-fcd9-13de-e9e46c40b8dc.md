
# PivotField.AddPageItem Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Adds an additional item to a multiple item page field.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AddPageItem**( **_Item_**,  **_ClearList_**)

 _expression_A variable that represents a  **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Item|Required| **String**| Source name of a **PivotItem** object, corresponding to the specific Online Analytical Processing (OLAP) member unique name.|
|ClearList|Optional| **Variant**|If  **False** (default), adds a page item to the existing list. If **True**, deletes all current items and adds Item.|

## Remarks
<a name="sectionSection1"> </a>

To avoid run-time errors, the data source must be an OLAP source, the field chosen must currently be in the page position, and the  ** [EnableMultiplePageItems](989fa662-cafb-00a1-effb-4a6c18327ea3.md)**property must be set to  **True**.


## Example
<a name="sectionSection2"> </a>

In this example, Microsoft Excel adds a page item with a source name titled "[Product].[All Products].[Food].[Eggs]". This example assumes an OLAP PivotTable exists on the active worksheet.


```
Sub UseAddPageItem() 
 
 ' The source is an OLAP database and you can manually reorder items. 
 ActiveSheet.PivotTables(1).CubeFields("[Product]"). _ 
 EnableMultiplePageItems = True 
 
 ' Add the page item titled "[Product].[All Products].[Food].[Eggs]". 
 ActiveSheet.PivotTables(1).PivotFields("[Product]").AddPageItem ( _ 
 "[Product].[All Products].[Food].[Eggs]") 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PivotField Object](52784960-e2da-b43a-1e37-2d4dae61c6d8.md)
#### Other resources


 [PivotField Object Members](4a6ea12a-072c-a386-c855-7bf5f6eadd46.md)


# Shape.RemoveCatalogMergeArea Method (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Deletes the catalog merge area from the specified publication page. All shapes contained in the catalog merge area remain in place on the page, but are no longer connected to the catalog merge data source.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RemoveCatalogMergeArea**

 _expression_A variable that represents a  **Shape** object.


## Remarks
<a name="sectionSection1"> </a>

Removing a catalog merge area from a publication page does not disconnect the data source from the publication. Use the  ** [IsDataSourceConnected](b62422ab-12f7-1151-d8d1-1cb32de18160.md)** property of the ** [Document](44f02255-ff5b-bcfe-900f-61c8fdf61ef3.md)** object to determine if a data source is connected to a publication.

Use the  ** [AddCatalogMergeArea](4af86b99-5a3a-b9f3-d269-16d635d35c83.md)** method of the ** [Shapes](52e069a6-d54b-a11a-1cba-96174329cb02.md)** collection to add a catalog merge area to a publication. A publication page can contain only one catalog merge area.


## Example
<a name="sectionSection2"> </a>

The following example tests whether any page in the specified publication contains a catalog merge area. If any page does, all the shapes are removed from the catalog merge area and deleted, and the catalog merge area is then removed from the publication.


```
Sub DeleteCatalogMergeAreaAndAllShapesWithin() 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 Dim strName As String 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 With mmLoop.CatalogMergeItems 
 For intCount = .Count To 1 Step -1 
 strName = mmLoop.CatalogMergeItems.Item(intCount).Name 
 .Item(intCount).RemoveFromCatalogMergeArea 
 pgPage.Shapes(strName).Delete 
 Next 
 End With 
 mmLoop.RemoveCatalogMergeArea 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
 End Sub 

```


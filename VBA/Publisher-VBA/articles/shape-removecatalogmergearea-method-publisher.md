---
title: Shape.RemoveCatalogMergeArea Method (Publisher)
keywords: vbapb10.chm5308691
f1_keywords:
- vbapb10.chm5308691
ms.prod: publisher
api_name:
- Publisher.Shape.RemoveCatalogMergeArea
ms.assetid: addff960-562e-b8e8-ec56-ddcf2b9ccaa7
ms.date: 06/08/2017
---


# Shape.RemoveCatalogMergeArea Method (Publisher)

Deletes the catalog merge area from the specified publication page. All shapes contained in the catalog merge area remain in place on the page, but are no longer connected to the catalog merge data source.


## Syntax

 _expression_. **RemoveCatalogMergeArea**

 _expression_A variable that represents a  **Shape** object.


## Remarks

Removing a catalog merge area from a publication page does not disconnect the data source from the publication. Use the  **[IsDataSourceConnected](document-isdatasourceconnected-property-publisher.md)** property of the **[Document](document-object-publisher.md)** object to determine if a data source is connected to a publication.

Use the  **[AddCatalogMergeArea](shapes-addcatalogmergearea-method-publisher.md)** method of the **[Shapes](shapes-object-publisher.md)** collection to add a catalog merge area to a publication. A publication page can contain only one catalog merge area.


## Example

The following example tests whether any page in the specified publication contains a catalog merge area. If any page does, all the shapes are removed from the catalog merge area and deleted, and the catalog merge area is then removed from the publication.


```vb
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



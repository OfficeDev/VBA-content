---
title: Shape.CatalogMergeItems Property (Publisher)
keywords: vbapb10.chm5308690
f1_keywords:
- vbapb10.chm5308690
ms.prod: publisher
api_name:
- Publisher.Shape.CatalogMergeItems
ms.assetid: 1dcf4ae0-7a18-f1d5-2176-1912c63eefcc
ms.date: 06/08/2017
---


# Shape.CatalogMergeItems Property (Publisher)

Returns a  **CatalogMergeShapes** collection that represents the shapes included in the catalog merge area. Read-only.


## Syntax

 _expression_. **CatalogMergeItems**

 _expression_A variable that represents a  **Shape** object.


### Return Value

CatalogMergeShapes


## Remarks

The catalog merge area can contain picture and text data fields you have inserted, in addition to other design elements you choose.


## Example

The following example tests whether any page in the specified publication contains a catalog merge area and, if it does, it returns a list of the shapes it contains.


```vb
Sub ListCatalogMergeAreaContents() 
 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 
 With mmLoop.CatalogMergeItems 
 For intCount = 1 To .Count 
 Debug.Print "Shape ID: " &; _ 
 mmLoop.CatalogMergeItems.Item(intCount).ID 
 Debug.Print "Shape Name: " &; _ 
 mmLoop.CatalogMergeItems.Item(intCount).Name 
 Next 
 End With 
 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
End Sub
```



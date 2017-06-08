---
title: CatalogMergeShapes Object (Publisher)
keywords: vbapb10.chm8454143
f1_keywords:
- vbapb10.chm8454143
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes
ms.assetid: 1108e9a4-57ef-2b1a-0998-54b6fad838da
ms.date: 06/08/2017
---


# CatalogMergeShapes Object (Publisher)

Represents the shapes contained in the catalog merge area of the specified publication.
 


## Remarks

The catalog merge area is automatically resized to accommodate objects that are larger then the merge area, or that are positioned outside the catalog merge area when they are added.
 

 
Shapes inside the catalog merge area are automatically resized or repositioned if the catalog merge area is decreased in size or moved.
 

 
The catalog merge area can contain picture and text data fields you have inserted, in addtion to other design elements you choose. 
 

 

## Example

Use the  **[CatalogMergeItems](shape-catalogmergeitems-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** objects to return the contents of the catalog merge area. The following example tests whether the specified publication contains a catalog merge area. If it does, it returns a list of the shapes it contains.
 

 

```
Sub ListCatalogMergeAreaContents() 
 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 
 With mmLoop.CatalogMergeItems 
 For intCount = 1 To .Count 
 Debug.Print "Shape ID: " &amp; _ 
 mmLoop.CatalogMergeItems.Item(intCount).ID 
 Debug.Print "Shape Name: " &amp; _ 
 mmLoop.CatalogMergeItems.Item(intCount).Name 
 Next 
 End With 
 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
End Sub 

```

Use the  **[AddToCatalogMergeArea](shape-addtocatalogmergearea-method-publisher.md)** method of the **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** objects to add shapes to a catalog merge area. The following example adds a rectangle to the catalog merge area in the specified publication. This example assumes a catalog merge area has been added to the first page of the publication.
 

 



```
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```

Use  **CatalogMergeItems** (index), where index is index number, to return a single catalog merge area shape. The following example removes the first shape from the catalog merge area.
 

 



```
ThisDocument.Pages(1).Shapes(1).CatalogMergeItems(1).RemoveFromCatalogMergeArea
```

Use the  **[RemoveFromCatalogMergeArea](shape-removefromcatalogmergearea-method-publisher.md)** method of the **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** objects to remove shapes from a catalog merge area. Removed shapes are not deleted, but are instead placed on the publication page containing the catalog merge area. The following example tests whether the specified publication contains a catalog merge area. If it does, all the shapes are removed from the catalog merge area and deleted, and the catalog merge area is then removed from the publication.
 

 



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


## Methods



|**Name**|
|:-----|
|[Item](catalogmergeshapes-item-method-publisher.md)|
|[Range](catalogmergeshapes-range-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](catalogmergeshapes-application-property-publisher.md)|
|[Count](catalogmergeshapes-count-property-publisher.md)|
|[HorizontalRepeat](catalogmergeshapes-horizontalrepeat-property-publisher.md)|
|[Parent](catalogmergeshapes-parent-property-publisher.md)|
|[VerticalRepeat](catalogmergeshapes-verticalrepeat-property-publisher.md)|


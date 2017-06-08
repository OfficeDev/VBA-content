---
title: CatalogMergeShapes.HorizontalRepeat Property (Publisher)
keywords: vbapb10.chm8388613
f1_keywords:
- vbapb10.chm8388613
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.HorizontalRepeat
ms.assetid: 1c3f1093-294f-e7b3-02ca-803ce7437d49
ms.date: 06/08/2017
---


# CatalogMergeShapes.HorizontalRepeat Property (Publisher)

Returns a  **Long** that represents the number of times the catalog merge area will repeat across the target publication page when the catalog merge is executed. Read-only.


## Syntax

 _expression_. **HorizontalRepeat**

 _expression_A variable that represents a  **CatalogMergeShapes** object.


### Return Value

Long


## Remarks

When the catalog merge is executed, the catalog merge area repeats once for each selected record in the specified data source.

The number of times the catalog merge area repeats across the page is determined by the width of the area. Use the  **[Width](shape-width-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object to return or set the horizontal size of the catalog merge area.

The  **[VerticalRepeat](catalogmergeshapes-verticalrepeat-property-publisher.md)** property of the **[CatalogMergeShapes](catalogmergeshapes-object-publisher.md)** object represents the number of times the catalog merge area repeats vertically down the target publication page.


## Example

The following example returns the number of times the catalog merge area will repeat horizontally and vertically on the target publication page when the catalog merge is performed. This example assumes the catalog merge area is the first shape on the first page of the specified publication.


```vb
Sub CatalogMergeDimensions() 
 
 With ThisDocument.Pages(1).Shapes(1) 
 Debug.Print .Width 
 Debug.Print .CatalogMergeItems.HorizontalRepeat 
 Debug.Print .Height 
 Debug.Print .CatalogMergeItems.VerticalRepeat 
 End With 
 
End Sub
```



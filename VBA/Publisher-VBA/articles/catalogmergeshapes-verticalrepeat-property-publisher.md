---
title: CatalogMergeShapes.VerticalRepeat Property (Publisher)
keywords: vbapb10.chm8388614
f1_keywords:
- vbapb10.chm8388614
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.VerticalRepeat
ms.assetid: 2a4852d6-14ee-7fa9-ea5e-170033c3a56d
ms.date: 06/08/2017
---


# CatalogMergeShapes.VerticalRepeat Property (Publisher)

Returns a  **Long** that represents the number of times the catalog merge area will repeat down the target publication page when the catalog merge is executed. Read-only.


## Syntax

 _expression_. **VerticalRepeat**

 _expression_A variable that represents a  **CatalogMergeShapes** object.


### Return Value

Long


## Remarks

When the catalog merge is executed, the catalog merge area repeats once for each selected record in the specified data source.

The number of times the catalog merge area repeats down the page is determined by the height of the area. Use the  **[Height](shape-height-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object to return or set the vertical size of the catalog merge area.

The  **[HorizontalRepeat](catalogmergeshapes-horizontalrepeat-property-publisher.md)** property of the **[CatalogMergeShapes](catalogmergeshapes-object-publisher.md)** object represents the number of times the catalog merge area repeats horizontally across the target publication page.


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



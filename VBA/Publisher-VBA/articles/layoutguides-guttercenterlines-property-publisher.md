---
title: LayoutGuides.GutterCenterlines Property (Publisher)
keywords: vbapb10.chm1114130
f1_keywords:
- vbapb10.chm1114130
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.GutterCenterlines
ms.assetid: 7a5b1aef-85c7-548f-15e9-2c3b7327b439
ms.date: 06/08/2017
---


# LayoutGuides.GutterCenterlines Property (Publisher)

Returns or sets a value that specifies whether to add a center line between the columns and rows of the gutter guides in a master page. Read/write  **Boolean**.


## Syntax

 _expression_. **GutterCenterlines**

 _expression_A variable that represents a  **LayoutGuides** object.


### Return Value

Boolean


## Remarks

The  **GutterCenterlines** property can only be used if the ** [LayoutGuides.Rows](layoutguides-rows-property-publisher.md)** property or the ** [LayoutGuides.Columns](layoutguides-columns-property-publisher.md)** property is greater than 1.

If  **True**, a red line appears in the center of the gutter guides. If  **False**, no line appears in the center of the gutter guides. The default value is  **False**.


## Example

The following example modifies the first master page of the active publication to have three rows, three columns, and red center lines drawn in the gutter guides. Any pages added to the publication after this point will have red center lines drawn in the gutter guides.


```vb
Dim theMasterPage As page 
Dim theLayoutGuides As LayoutGuides 
 
Set theMasterPage = ActiveDocument.MasterPages(1) 
Set theLayoutGuides = theMasterPage.LayoutGuides 
 
With theLayoutGuides 
 .Rows = 3 
 .Columns = 3 
 .GutterCenterlines = True 
End With
```



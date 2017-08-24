---
title: Page.LayoutGuides Property (Publisher)
keywords: vbapb10.chm393270
f1_keywords:
- vbapb10.chm393270
ms.prod: publisher
api_name:
- Publisher.Page.LayoutGuides
ms.assetid: eb9ac463-2b9f-9c68-b58f-6d93fe4993c8
ms.date: 06/08/2017
---


# Page.LayoutGuides Property (Publisher)

Returns a  **[LayoutGuides](layoutguides-object-publisher.md)** object consisting of the margin and grid layout guides for all pages including master pages in the publication.


## Syntax

 _expression_. **LayoutGuides**

 _expression_A variable that represents a  **Page** object.


## Example

The following example changes the grid layout guides so that there are three columns and five rows.


```vb
Dim layTemp As LayoutGuides 
 
Set layTemp = ActiveDocument.LayoutGuides 
 
With layTemp 
 .Rows = 5 
 .Columns = 3 
End With 

```



---
title: Document.Tags Property (Publisher)
keywords: vbapb10.chm196661
f1_keywords:
- vbapb10.chm196661
ms.prod: publisher
api_name:
- Publisher.Document.Tags
ms.assetid: d8baaf50-86ad-1997-c1b3-e54a77a3ee5b
ms.date: 06/08/2017
---


# Document.Tags Property (Publisher)

Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.


## Syntax

 _expression_. **Tags**

 _expression_A variable that represents a  **Document** object.


## Example

This example adds a tag to each oval shape on the first page of the active publication.


```vb
Dim shp As Shape 
Dim tagsAll As Tags 
Dim tagLoop As Tag 
Dim blnTag As Boolean 
 
With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.AutoShapeType = msoShapeOval Then 
 Set tagsAll = shp.Tags 
 blnTag = False 
 
 For Each tagLoop In tagsAll 
 If tagLoop.Name = "Shape" Then 
 blnTag = True 
 Exit For 
 End If 
 Next tagLoop 
 
 If blnTag = False Then 
 tagsAll.Add Name:="Shape", Value:="Oval" 
 End If 
 End If 
 Next shp 
End With 

```



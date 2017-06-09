---
title: Page.Tags Property (Publisher)
keywords: vbapb10.chm393235
f1_keywords:
- vbapb10.chm393235
ms.prod: publisher
api_name:
- Publisher.Page.Tags
ms.assetid: 94a8be36-20c2-65bc-b1e2-41f24703b264
ms.date: 06/08/2017
---


# Page.Tags Property (Publisher)

Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.


## Syntax

 _expression_. **Tags**

 _expression_A variable that represents a  **Page** object.


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



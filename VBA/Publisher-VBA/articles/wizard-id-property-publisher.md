---
title: Wizard.ID Property (Publisher)
keywords: vbapb10.chm1441795
f1_keywords:
- vbapb10.chm1441795
ms.prod: publisher
api_name:
- Publisher.Wizard.ID
ms.assetid: ce7df9d3-052a-6cb6-e24d-4cb5192551d0
ms.date: 06/08/2017
---


# Wizard.ID Property (Publisher)

Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **Wizard** object.


## Example

This example displays the type for each shape on the first page of the active publication.


```vb
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```



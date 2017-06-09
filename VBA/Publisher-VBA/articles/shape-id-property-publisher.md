---
title: Shape.ID Property (Publisher)
keywords: vbapb10.chm2228325
f1_keywords:
- vbapb10.chm2228325
ms.prod: publisher
api_name:
- Publisher.Shape.ID
ms.assetid: df4ccd93-e3fa-eeef-b5ea-e99aa0dde199
ms.date: 06/08/2017
---


# Shape.ID Property (Publisher)

Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **Shape** object.


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



---
title: ShapeRange.ID Property (Publisher)
keywords: vbapb10.chm2293861
f1_keywords:
- vbapb10.chm2293861
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ID
ms.assetid: d7ad646b-be40-2ac4-9d3e-faa37f8bf456
ms.date: 06/08/2017
---


# ShapeRange.ID Property (Publisher)

Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **ShapeRange** object.


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



---
title: WizardProperty.ID Property (Publisher)
keywords: vbapb10.chm1572867
f1_keywords:
- vbapb10.chm1572867
ms.prod: publisher
api_name:
- Publisher.WizardProperty.ID
ms.assetid: 2827af5d-d002-029b-7f93-26befe459229
ms.date: 06/08/2017
---


# WizardProperty.ID Property (Publisher)

Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **WizardProperty** object.


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



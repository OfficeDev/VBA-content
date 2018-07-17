---
title: ShapeRange.OLEFormat Property (Publisher)
keywords: vbapb10.chm2293863
f1_keywords:
- vbapb10.chm2293863
ms.prod: publisher
api_name:
- Publisher.ShapeRange.OLEFormat
ms.assetid: 237b51e8-dced-3e21-d257-410121107a63
ms.date: 06/08/2017
---


# ShapeRange.OLEFormat Property (Publisher)

Returns an  **[OLEFormat](oleformat-object-publisher.md)** object that contains OLE formatting properties for the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent OLE objects.


## Syntax

 _expression_. **OLEFormat**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example loops through all the shapes on the first page of the active document and automatically updates all linked Excel worksheets.


```vb
Sub UpdateLinkedExcelSpreadsheets() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = msoLinkedOLEObject Then 
 If shp.OLEFormat.ProgId = "Excel.Sheet" Then 
 shp.LinkFormat.Update 
 End If 
 End If 
 Next shp 
End Sub
```



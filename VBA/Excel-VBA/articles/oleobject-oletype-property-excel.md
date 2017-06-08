---
title: OLEObject.OLEType Property (Excel)
keywords: vbaxl10.chm417077
f1_keywords:
- vbaxl10.chm417077
ms.prod: excel
api_name:
- Excel.OLEObject.OLEType
ms.assetid: ec26dc86-5a31-ca8b-97c7-fe513fb283b1
ms.date: 06/08/2017
---


# OLEObject.OLEType Property (Excel)

Returns the OLE object type. Can be one of the following  **XlOLEType** constants: **xlOLELink** or **xlOLEEmbed** . Returns **xlOLELink** if the object is linked (it exists outside of the file), or returns **xlOLEEmbed** if the object is embedded (it's entirely contained within the file). Read-only **Long** .


## Syntax

 _expression_ . **OLEType**

 _expression_ A variable that represents an **OLEObject** object.


## Example

This example creates a list of link types for OLE objects on Sheet1. The list appears on a new worksheet created by the example.


```vb
Set newSheet = Worksheets.Add 
i = 2 
newSheet.Range("A1").Value = "Name" 
newSheet.Range("B1").Value = "Link Type" 
For Each obj In Worksheets("Sheet1").OLEObjects 
 newSheet.Cells(i, 1).Value = obj.Name 
 If obj.OLEType = xlOLELink Then 
 newSheet.Cells(i, 2) = "Linked" 
 Else 
 newSheet.Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)


---
title: OLEObject.progID Property (Excel)
keywords: vbaxl10.chm417083
f1_keywords:
- vbaxl10.chm417083
ms.prod: excel
api_name:
- Excel.OLEObject.progID
ms.assetid: cbec1e95-6bdd-ce55-f426-28dcf4191897
ms.date: 06/08/2017
---


# OLEObject.progID Property (Excel)

Returns the programmatic identifiers for the object. Read-only  **String** .


## Syntax

 _expression_ . **progID**

 _expression_ A variable that represents an **OLEObject** object.


## Example

This example creates a list of the programmatic identifiers for the OLE objects on worksheet one.


```vb
rw = 0 
For Each o in Worksheets(1).OLEObjects 
 With Worksheets(2) 
 rw = rw + 1 
 .cells(rw, 1).Value = o.ProgId 
 End With 
Next
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)


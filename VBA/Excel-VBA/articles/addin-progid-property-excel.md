---
title: AddIn.progID Property (Excel)
keywords: vbaxl10.chm185082
f1_keywords:
- vbaxl10.chm185082
ms.prod: excel
api_name:
- Excel.AddIn.progID
ms.assetid: a1c1da74-119a-514e-cb5e-77981299b84d
ms.date: 06/08/2017
---


# AddIn.progID Property (Excel)

Returns the programmatic identifiers for the object. Read-only  **String** .


## Syntax

 _expression_ . **progID**

 _expression_ A variable that represents an **AddIn** object.


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


[AddIn Object](addin-object-excel.md)


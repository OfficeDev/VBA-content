---
title: OLEFormat.progID Property (Excel)
keywords: vbaxl10.chm632075
f1_keywords:
- vbaxl10.chm632075
ms.prod: excel
api_name:
- Excel.OLEFormat.progID
ms.assetid: 77156cae-46fc-2068-4dce-cb584e56b496
ms.date: 06/08/2017
---


# OLEFormat.progID Property (Excel)

Returns the programmatic identifiers for the object. Read-only  **String** .


## Syntax

 _expression_ . **progID**

 _expression_ A variable that represents an **OLEFormat** object.


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


[OLEFormat Object](oleformat-object-excel.md)


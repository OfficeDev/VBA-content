---
title: OLEDBError.Number Property (Excel)
keywords: vbaxl10.chm654076
f1_keywords:
- vbaxl10.chm654076
ms.prod: excel
api_name:
- Excel.OLEDBError.Number
ms.assetid: 9e88a0bb-1cbf-d98e-52a9-a8f9a0bde81c
ms.date: 06/08/2017
---


# OLEDBError.Number Property (Excel)

Returns a numeric value that specifies an error. The error number corresponds to a unique trap number corresponding to an error condition that resulted after the most recent OLE DB query. Read-only  **Long** .


## Syntax

 _expression_ . **Number**

 _expression_ A variable that represents an **OLEDBError** object.


## Example

This example displays the error number and other error information returned by the most recent OLE DB query.


```vb
Set objEr = Application.OLEDBErrors(1) 
MsgBox "The following error occurred:" &; _ 
 objEr.Number &; ", " &; objEr.Native &; ", " &; _ 
 objEr.ErrorString &; " : " &; objEr.SqlState
```


## See also


#### Concepts


[OLEDBError Object](oledberror-object-excel.md)


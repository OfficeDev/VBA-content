---
title: OLEDBError.Native Property (Excel)
keywords: vbaxl10.chm654075
f1_keywords:
- vbaxl10.chm654075
ms.prod: excel
api_name:
- Excel.OLEDBError.Native
ms.assetid: 2eae623f-7803-b3ce-467b-ee4f9c5c8c20
ms.date: 06/08/2017
---


# OLEDBError.Native Property (Excel)

Returns a provider-specific numeric value that specifies an error. The error number corresponds to an error condition that resulted after the most recent OLE DB query. Read-only  **Long** .


## Syntax

 _expression_ . **Native**

 _expression_ A variable that represents an **OLEDBError** object.


## Example

This example displays the native error number and other error information returned by the most recent OLE DB query.


```vb
Set objEr = Application.OLEDBErrors(1) 
MsgBox "The following error occurred:" &; _ 
 objEr.Number &; ", " &; objEr.Native &; ", " &; _ 
 objEr.ErrorString &; " : " &; objEr.SqlState
```


## See also


#### Concepts


[OLEDBError Object](oledberror-object-excel.md)


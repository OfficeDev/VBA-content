---
title: Application.OLEDBErrors Property (Excel)
keywords: vbaxl10.chm133244
f1_keywords:
- vbaxl10.chm133244
ms.prod: excel
api_name:
- Excel.Application.OLEDBErrors
ms.assetid: 0a42417f-f8b6-10bf-712a-44c1107f0f3e
ms.date: 06/08/2017
---


# Application.OLEDBErrors Property (Excel)

Returns the  **[OLEDBErrors](oledberrors-object-excel.md)** collection, which represents the error information returned by the most recent OLE DB query. Read-only.


## Syntax

 _expression_ . **OLEDBErrors**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays the error description and  **SqlState** property value for an OLE DB error returned by the most recent OLE DB query.


```vb
Set objEr = Application.OLEDBErrors.Item(1) 
MsgBox "The following error occurred:" &; _ 
 objEr.ErrorString &; " : " &; objEr.SqlState
```


## See also


#### Concepts


[Application Object](application-object-excel.md)


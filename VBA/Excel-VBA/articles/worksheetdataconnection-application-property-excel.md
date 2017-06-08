---
title: WorksheetDataConnection.Application Property (Excel)
keywords: vbaxl10.chm923073
f1_keywords:
- vbaxl10.chm923073
ms.prod: excel
ms.assetid: 79545289-efa9-ce0b-3268-4f73c410fb55
ms.date: 06/08/2017
---


# WorksheetDataConnection.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[WorksheetDataConnection Object (Excel)](worksheetdataconnection-object-excel.md) object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**


## See also


#### Other resources



[WorksheetDataConnection Object](worksheetdataconnection-object-excel.md)


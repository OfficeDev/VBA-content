---
title: ModelConnection.Application Property (Excel)
keywords: vbaxl10.chm921073
f1_keywords:
- vbaxl10.chm921073
ms.prod: excel
ms.assetid: 6d0ff59e-4d5d-c06c-4af8-33a69739f9e1
ms.date: 06/08/2017
---


# ModelConnection.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelConnection Object (Excel)](modelconnection-object-excel.md) object.


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



[ModelConnection Object](modelconnection-object-excel.md)


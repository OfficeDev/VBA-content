---
title: ModelTableNames.Application Property (Excel)
keywords: vbaxl10.chm953073
f1_keywords:
- vbaxl10.chm953073
ms.prod: excel
ms.assetid: 65e3c7a7-2cce-b9e3-cf72-6fc533ec9f83
ms.date: 06/08/2017
---


# ModelTableNames.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelTableNames Object (Excel)](modeltablenames-object-excel.md) object.


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



[ModelTableNames Object](modeltablenames-object-excel.md)


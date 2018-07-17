---
title: ModelTable.Application Property (Excel)
keywords: vbaxl10.chm933073
f1_keywords:
- vbaxl10.chm933073
ms.prod: excel
ms.assetid: c2138114-e623-f141-090d-22644f8d2477
ms.date: 06/08/2017
---


# ModelTable.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelTable Object (Excel)](modeltable-object-excel.md) object.


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



[ModelTable Object](modeltable-object-excel.md)


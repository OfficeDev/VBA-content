---
title: Application.DisplayClipboardWindow Property (Excel)
keywords: vbaxl10.chm133093
f1_keywords:
- vbaxl10.chm133093
ms.prod: excel
api_name:
- Excel.Application.DisplayClipboardWindow
ms.assetid: 16686caf-39ed-90fa-4a61-92b3f825cc6c
ms.date: 06/08/2017
---


# Application.DisplayClipboardWindow Property (Excel)

Returns  **True** if the Microsoft Office Clipboard can be displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayClipboardWindow**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel determines if the Office Clipboard can be displayed and notifies the user.


```vb
Sub SeeClipboard() 
 
 ' Determine if Office Clipboard can be displayed. 
 If Application.DisplayClipboardWindow = True Then 
 MsgBox "Office Clipboard can be displayed." 
 Else 
 MsgBox "Office Clipboard cannot be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)


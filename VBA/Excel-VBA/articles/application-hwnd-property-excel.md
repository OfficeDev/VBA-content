---
title: Application.Hwnd Property (Excel)
keywords: vbaxl10.chm133277
f1_keywords:
- vbaxl10.chm133277
ms.prod: excel
api_name:
- Excel.Application.Hwnd
ms.assetid: ed98b59c-1ebf-f319-f986-3406e4fdb766
ms.date: 06/08/2017
---


# Application.Hwnd Property (Excel)

Returns a  **Long** indicating the top-level window handle of the Microsoft Excel window. Read-only.


## Syntax

 _expression_ . **Hwnd**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel notifies the user of the top-level window handle of the Excel window.


```vb
Sub CheckHwnd() 
 
 MsgBox "The top-level window handle is: " &; _ 
 Application.Hwnd 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)


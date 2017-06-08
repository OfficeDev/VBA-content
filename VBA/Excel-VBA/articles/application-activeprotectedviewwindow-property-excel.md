---
title: Application.ActiveProtectedViewWindow Property (Excel)
keywords: vbaxl10.chm133331
f1_keywords:
- vbaxl10.chm133331
ms.prod: excel
api_name:
- Excel.Application.ActiveProtectedViewWindow
ms.assetid: 2202c3b4-8880-7a26-8a56-8f2d2e7b7343
ms.date: 06/08/2017
---


# Application.ActiveProtectedViewWindow Property (Excel)

Returns a  **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object that represents the active **Protected View** window (the window on top). Read-only. Returns **Nothing** if there are no **Protected View** windows open. Read-only


## Syntax

 _expression_ . **ActiveProtectedViewWindow**

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


## Example

The following code example displays the name ( **Caption** property) of the active **Protected View** window.


```vb
MsgBox "The name of the active Protected View window is " &; ActiveProtectedWindow.Caption
```


## See also


#### Concepts


[Application Object](application-object-excel.md)


---
title: ProtectedViewWindow.Caption Property (Excel)
keywords: vbaxl10.chm914074
f1_keywords:
- vbaxl10.chm914074
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Caption
ms.assetid: fe3f8026-71e2-3a5a-9376-7b9d93f97b6f
ms.date: 06/08/2017
---


# ProtectedViewWindow.Caption Property (Excel)

Returns or sets a  **Variant** value that represents the name that appears in the title bar of the **Protected View** window. Read/write


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


## Example

The following code example displays the name ( **Caption** property) of the active **Protected View** window.


```vb
MsgBox "The name of the active Protected View window is " &; ActiveProtectedWindow.Caption
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)


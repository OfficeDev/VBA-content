---
title: Application.WindowState Property (Excel)
keywords: vbaxl10.chm133234
f1_keywords:
- vbaxl10.chm133234
ms.prod: excel
api_name:
- Excel.Application.WindowState
ms.assetid: f53d2bb8-b862-c55f-d9d5-68e705ca3415
ms.date: 06/08/2017
---


# Application.WindowState Property (Excel)

Returns or sets the state of the window. Read/write  **[XlWindowState](xlwindowstate-enumeration-excel.md)** .


## Syntax

 _expression_ . **WindowState**

 _expression_ A variable that represents an **Application** object.


## Example

This example maximizes the application window in Microsoft Excel.


```vb
Application.WindowState = xlMaximized
```

This example expands the active window to the maximum size available (assuming that the window isn't already maximized).




```vb
With ActiveWindow 
 .WindowState = xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight 
 .Width = Application.UsableWidth 
End With 

```


## See also


#### Concepts


[Application Object](application-object-excel.md)


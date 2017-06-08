---
title: Window.UsableHeight Property (Excel)
keywords: vbaxl10.chm356119
f1_keywords:
- vbaxl10.chm356119
ms.prod: excel
api_name:
- Excel.Window.UsableHeight
ms.assetid: e1cbcaa1-779a-1757-0a95-9e53e374ef7c
ms.date: 06/08/2017
---


# Window.UsableHeight Property (Excel)

Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only  **Double** .


## Syntax

 _expression_ . **UsableHeight**

 _expression_ A variable that represents a **Window** object.


## Example

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


[Window Object](window-object-excel.md)


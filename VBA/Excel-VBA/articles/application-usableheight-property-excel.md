---
title: Application.UsableHeight Property (Excel)
keywords: vbaxl10.chm133222
f1_keywords:
- vbaxl10.chm133222
ms.prod: excel
api_name:
- Excel.Application.UsableHeight
ms.assetid: 536d2d03-0ce8-c28a-5a94-461fcfbd4ebf
ms.date: 06/08/2017
---


# Application.UsableHeight Property (Excel)

Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only  **Double** .


## Syntax

 _expression_ . **UsableHeight**

 _expression_ A variable that represents an **Application** object.


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


[Application Object](application-object-excel.md)


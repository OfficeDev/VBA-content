---
title: Application.Width Property (Excel)
keywords: vbaxl10.chm133232
f1_keywords:
- vbaxl10.chm133232
ms.prod: excel
api_name:
- Excel.Application.Width
ms.assetid: eeb8ff27-d219-bade-3e0b-aed6e34d17d7
ms.date: 06/08/2017
---


# Application.Width Property (Excel)

Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the application window to its right edge.


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents an **Application** object.


## Remarks

 If the window is minimized, **Width** is read-only and returns the width of the window icon.


## Example

This example expands the active window to the maximum size available (assuming that the window isn't maximized).


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


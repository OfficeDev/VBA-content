---
title: Application.ActiveWindow Property (PowerPoint)
keywords: vbapp10.chm502004
f1_keywords:
- vbapp10.chm502004
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActiveWindow
ms.assetid: 762c1c6a-1f8a-f47a-7b75-006c745caee0
ms.date: 06/08/2017
---


# Application.ActiveWindow Property (PowerPoint)

Returns a  **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the active document window. Read-only.


## Syntax

 _expression_. **ActiveWindow**

 _expression_ A variable that represents an **Application** object.


### Return Value

DocumentWindow


## Example

This example minimizes the active window.


```vb
Application.ActiveWindow.WindowState = ppWindowMinimized
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)


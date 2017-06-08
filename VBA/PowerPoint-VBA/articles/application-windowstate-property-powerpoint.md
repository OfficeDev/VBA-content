---
title: Application.WindowState Property (PowerPoint)
keywords: vbapp10.chm502030
f1_keywords:
- vbapp10.chm502030
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowState
ms.assetid: 128f7da4-3cc3-1cda-6298-8bbc0b39a25c
ms.date: 06/08/2017
---


# Application.WindowState Property (PowerPoint)

Returns or sets the state of the specified window. Read/write.


## Syntax

 _expression_. **WindowState**

 _expression_ A variable that represents an **Application** object.


### Return Value

PpWindowState


## Remarks

The value of the  **WindowState** property can be one of these **PpWindowState** constants.


||
|:-----|
|**ppWindowMaximized**|
|**ppWindowMinimized**|
|**ppWindowNormal**|
When the state of the window is  **ppWindowNormal**, the window is neither maximized nor minimized.


## Example

This example maximizes the active window.


```vb
Application.ActiveWindow.WindowState = ppWindowMaximized
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)


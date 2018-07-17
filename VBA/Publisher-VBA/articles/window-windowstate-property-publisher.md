---
title: Window.WindowState Property (Publisher)
keywords: vbapb10.chm262160
f1_keywords:
- vbapb10.chm262160
ms.prod: publisher
api_name:
- Publisher.Window.WindowState
ms.assetid: 063ede5e-f279-09e3-5672-b634c752b927
ms.date: 06/08/2017
---


# Window.WindowState Property (Publisher)

Returns or sets a  **PbWindowState** constant indicating the state of the Microsoft Publisher window. Read/write.


## Syntax

 _expression_. **WindowState**

 _expression_A variable that represents a  **Window** object.


### Return Value

PbWindowState


## Remarks

The  **WindowState** property value can be one of these **PbWindowState** constants.



| **pbWindowStateMaximize**|
| **pbWindowStateMinimize**|
| **pbWindowStateNormal**|
When the state of the window is  **pbWindowStateNormal**, the window is neither maximized nor minimized.


## Example

This example maximizes the Publisher window.


```vb
ActiveWindow.WindowState = pbWindowStateMaximized
```



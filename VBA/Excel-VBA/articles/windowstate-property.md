---
title: WindowState Property
keywords: vbagr10.chm65932
f1_keywords:
- vbagr10.chm65932
ms.prod: excel
api_name:
- Excel.WindowState
ms.assetid: 22ce1105-6f4e-54d2-4f9a-216019462f04
ms.date: 06/08/2017
---


# WindowState Property

Returns or sets the state of the window. Read/write XlWindowState .



|XlWindowState can be one of these XlWindowState constants.|
| **xlMaximized**|
| **xlNormal**|
| **xlMinimized**|

 _expression_. **WindowState**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example maximizes the Microsoft Graph application window.


```
myChart.Application.WindowState = xlMaximized
```



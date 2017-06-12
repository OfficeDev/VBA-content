---
title: Application.hWndAccessApp Method (Access)
keywords: vbaac10.chm12552
f1_keywords:
- vbaac10.chm12552
ms.prod: access
api_name:
- Access.Application.hWndAccessApp
ms.assetid: 7a4f162a-e2de-728b-09e0-f9272ad52053
ms.date: 06/08/2017
---


# Application.hWndAccessApp Method (Access)

You can use the  **hWndAccessApp** method to determine the handle assigned by Microsoft Windows to the main Microsoft Access window.


## Syntax

 _expression_. **hWndAccessApp**

 _expression_ A variable that represents an **Application** object.


### Return Value

Long


## Remarks

The  **hWndAccessApp** method returns a **Long Integer** value set by Microsoft Access and is read-only.

You can use this method by using [Visual Basic](set-properties-by-using-visual-basic.md)when making calls to Windows application programming interface (API) functions or other external procedures that require a window handle as an argument.

To get the handle to a window containing a Microsoft Access object such as a Form or Report, use the  **hWnd** property.


## See also


#### Concepts


[Application Object](application-object-access.md)


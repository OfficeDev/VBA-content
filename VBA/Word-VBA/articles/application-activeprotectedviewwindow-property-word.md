---
title: Application.ActiveProtectedViewWindow Property (Word)
keywords: vbawd10.chm158335467
f1_keywords:
- vbawd10.chm158335467
ms.prod: word
api_name:
- Word.Application.ActiveProtectedViewWindow
ms.assetid: 2ba10f3d-3f43-5628-a5fc-3c65b290ef72
ms.date: 06/08/2017
---


# Application.ActiveProtectedViewWindow Property (Word)

Returns a [ProtectedViewWindow](protectedviewwindow-object-word.md) object that represents the active protected view window. Read-only.


## Syntax

 _expression_ . **ActiveProtectedViewWindow**

 _expression_ An expression that returns a **Application** object.


## Example

The following code example displays the caption text for the active protected view window.


```vb
MsgBox ActiveProtectedViewWindow.Caption 

```


## See also


#### Concepts


[Application Object](application-object-word.md)


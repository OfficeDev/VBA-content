---
title: ProtectedViewWindow.Top Property (Word)
keywords: vbawd10.chm231735299
f1_keywords:
- vbawd10.chm231735299
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Top
ms.assetid: 3acaef1b-11a8-9f22-3841-049ae9e2ecd3
ms.date: 06/08/2017
---


# ProtectedViewWindow.Top Property (Word)

Returns or sets the vertical position, in points, of the specified protected view window. Read/write  **Long**


## Syntax

 _expression_ . **Top**

 _expression_ An expression that returns a **ProtectedViewWindow** object.


## Example

The following code example sets the vertical position of the active protected view window to 100 point.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Left = 0 
 .Top = 100 
End With 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)


---
title: ProtectedViewWindow.Left Property (Word)
keywords: vbawd10.chm231735298
f1_keywords:
- vbawd10.chm231735298
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Left
ms.assetid: 55ca42b8-bed4-3b7e-fd0b-66dc2ea936c3
ms.date: 06/08/2017
---


# ProtectedViewWindow.Left Property (Word)

Returns or sets a  **Long** , in points, that represents the horizontal position of the specified protected view window. Read/write.


## Syntax

 _expression_ . **Left**

 _expression_ An expression that returns a **[ProtectedViewWindow](protectedviewwindow-object-word.md)** object.


## Example

The following code example sets the horizontal position of the active protected view window to 100 point.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Left = 100 
 .Top = 0 
End With
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)


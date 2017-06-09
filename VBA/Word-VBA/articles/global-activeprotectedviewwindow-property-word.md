---
title: Global.ActiveProtectedViewWindow Property (Word)
keywords: vbawd10.chm163119219
f1_keywords:
- vbawd10.chm163119219
ms.prod: word
api_name:
- Word.Global.ActiveProtectedViewWindow
ms.assetid: 4023444a-f433-7f38-bbc8-6055ed03cb6a
ms.date: 06/08/2017
---


# Global.ActiveProtectedViewWindow Property (Word)

Returns a [ProtectedViewWindow](protectedviewwindow-object-word.md) object that represents the active protected view window (the protected view window with the focus). Read-only.


## Syntax

 _expression_ . **ActiveProtectedViewWindow**

 _expression_ An expression that returns a **Global** object.


## Remarks

If there are no windows open, using this property causes an error.


## Example

The following code example displays the caption text for the active protected view window.


```vb
Sub WindowCaption() 
 MsgBox ActiveProtectedViewWindow.Caption 
End Sub
```


## See also


#### Concepts


[Global Object](global-object-word.md)


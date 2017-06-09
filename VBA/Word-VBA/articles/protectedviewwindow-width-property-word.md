---
title: ProtectedViewWindow.Width Property (Word)
keywords: vbawd10.chm231735300
f1_keywords:
- vbawd10.chm231735300
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Width
ms.assetid: 607ec503-2096-4b4a-fce5-9979bea6c847
ms.date: 06/08/2017
---


# ProtectedViewWindow.Width Property (Word)

Returns or sets the width, in points, of the specified protectd view window. Read/write  **Long** .


## Syntax

 _expression_ . **Width**

 _expression_ An expression that returns a **[ProtectedViewWindow](protectedviewwindow-object-word.md)** object.


## Example

The following code example changes the state, height, and width of the active protected view window.


```vb
With ActiveProtectedViewWindow 
 .WindowState = wdWindowStateNormal 
 .Height = 400 
 .Width = 500 
End With 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)


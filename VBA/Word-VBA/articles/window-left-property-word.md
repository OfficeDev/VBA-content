---
title: Window.Left Property (Word)
keywords: vbawd10.chm157417477
f1_keywords:
- vbawd10.chm157417477
ms.prod: word
api_name:
- Word.Window.Left
ms.assetid: 915fe24c-084b-f7f0-46ad-a69c186cf737
ms.date: 06/08/2017
---


# Window.Left Property (Word)

Returns or sets a  **Long** that represents the horizontal position of the specified window, measured in points. Read/write.


## Syntax

 _expression_ . **Left**

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example sets the horizontal position of the active window to 100 points.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Left = 100 
 .Top = 0 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)


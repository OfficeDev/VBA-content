---
title: Window.DisplayHorizontalScrollBar Property (Word)
keywords: vbawd10.chm157417492
f1_keywords:
- vbawd10.chm157417492
ms.prod: word
api_name:
- Word.Window.DisplayHorizontalScrollBar
ms.assetid: c52d2cc8-d7ce-0b95-e97c-e41e449e4be6
ms.date: 06/08/2017
---


# Window.DisplayHorizontalScrollBar Property (Word)

 **True** if a horizontal scroll bar is displayed for the specified window. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayHorizontalScrollBar**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example displays vertical and horizontal scroll bars for the active window.


```vb
With ActiveDocument.ActiveWindow 
 .DisplayHorizontalScrollBar = True 
 .DisplayVerticalScrollBar = True 
End With
```

This example toggles the horizontal scroll bar of the window for Document1.




```vb
Dim winTemp As Window 
 
Set winTemp = Windows("Document1") 
 
winTemp.DisplayHorizontalScrollBar = _ 
 Not winTemp.DisplayHorizontalScrollBar
```


## See also


#### Concepts


[Window Object](window-object-word.md)


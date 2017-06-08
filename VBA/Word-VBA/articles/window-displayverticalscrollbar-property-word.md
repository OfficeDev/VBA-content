---
title: Window.DisplayVerticalScrollBar Property (Word)
keywords: vbawd10.chm157417491
f1_keywords:
- vbawd10.chm157417491
ms.prod: word
api_name:
- Word.Window.DisplayVerticalScrollBar
ms.assetid: bac2fcd6-d9b9-e922-b4ac-c891de68f6f3
ms.date: 06/08/2017
---


# Window.DisplayVerticalScrollBar Property (Word)

 **True** if a vertical scroll bar is displayed for the specified window. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayVerticalScrollBar**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example displays the vertical and horizontal scroll bars for each window in the Windows collection.


```vb
Dim winLoop As Window 
 
For Each winLoop In Windows 
 winLoop.DisplayVerticalScrollBar = True 
 winLoop.DisplayHorizontalScrollBar = True 
Next winLoop
```

This example toggles the vertical scroll bar for the active window.




```vb
Dim winTemp As Window 
 
Set winTemp = ActiveDocument.ActiveWindow 
winTemp.DisplayVerticalScrollBar = _ 
 Not winTemp.DisplayVerticalScrollBar
```


## See also


#### Concepts


[Window Object](window-object-word.md)


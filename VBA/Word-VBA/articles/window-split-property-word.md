---
title: Window.Split Property (Word)
keywords: vbawd10.chm157417481
f1_keywords:
- vbawd10.chm157417481
ms.prod: word
api_name:
- Word.Window.Split
ms.assetid: 97631d2f-577f-1a19-18e9-ae0ba92da054
ms.date: 06/08/2017
---


# Window.Split Property (Word)

 **True** if the window is split into multiple panes. Read/write **Boolean** .


## Syntax

 _expression_ . **Split**

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example splits the active window into two equal-sized window panes.


```vb
ActiveDocument.ActiveWindow.Split = True
```

If the Document1 window is split, this example closes the active pane.




```vb
If Windows("Document1").Split = True Then 
 Windows("Document1").ActivePane.Close 
End If
```


## See also


#### Concepts


[Window Object](window-object-word.md)


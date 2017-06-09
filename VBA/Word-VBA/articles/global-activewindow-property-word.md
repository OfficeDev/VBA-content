---
title: Global.ActiveWindow Property (Word)
keywords: vbawd10.chm163119108
f1_keywords:
- vbawd10.chm163119108
ms.prod: word
api_name:
- Word.Global.ActiveWindow
ms.assetid: 645913c3-0724-1604-9ac0-4e1b4e81439d
ms.date: 06/08/2017
---


# Global.ActiveWindow Property (Word)

Returns a  **[Window](window-object-word.md)** object that represents the active window (the window with the focus). Read-only.


## Syntax

 _expression_ . **ActiveWindow**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

If there are no windows open, using this property causes an error. 


## Example

This example displays the caption text for the active window.


```vb
Sub WindowCaption() 
 MsgBox ActiveDocument.ActiveWindow.Caption 
End Sub
```

This example opens a new window for the active window of the active document and then tiles all the windows.




```vb
Sub WindowTiled() 
 Dim wndTileWindow As Window 
 
 Set wndTileWindow = ActiveDocument.ActiveWindow.NewWindow 
 Windows.Arrange ArrangeStyle:=wdTiled 
End Sub
```

This example splits the first document window.




```vb
Sub WindowSplit() 
 Documents(1).ActiveWindow.Split = True 
End Sub
```


## See also


#### Concepts


[Global Object](global-object-word.md)


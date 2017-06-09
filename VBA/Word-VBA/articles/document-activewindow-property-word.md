---
title: Document.ActiveWindow Property (Word)
keywords: vbawd10.chm158007338
f1_keywords:
- vbawd10.chm158007338
ms.prod: word
api_name:
- Word.Document.ActiveWindow
ms.assetid: 707fe9e8-16de-c4aa-a0f7-6a4570d16cdd
ms.date: 06/08/2017
---


# Document.ActiveWindow Property (Word)

Returns a  **[Window](window-object-word.md)** object that represents the active window (the window with the focus). Read-only.


## Syntax

 _expression_ . **ActiveWindow**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If there are no windows open, using the  **ActiveWindow** property generates an error occurs.


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


[Document Object](document-object-word.md)


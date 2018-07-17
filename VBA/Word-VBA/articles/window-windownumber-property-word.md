---
title: Window.WindowNumber Property (Word)
keywords: vbawd10.chm157417490
f1_keywords:
- vbawd10.chm157417490
ms.prod: word
api_name:
- Word.WindowNumber
ms.assetid: 9fe66956-664f-083e-62fe-7c2919619615
ms.date: 06/08/2017
---


# Window.WindowNumber Property (Word)

Returns the window number of the document displayed in the specified window. For example, if the caption of the window is "Sales.doc:2", this property returns the number 2. Read-only  **Long** .


## Syntax

 _expression_ . **WindowNumber**

 _expression_ An expression that returns a **[Window](window-object-word.md)** object.


## Remarks

Use the property to return the number of the specified window in the  **[Windows](windows-object-word.md)** collection.


## Example

This example retrieves the window number of the active window, opens a new window, and then activates the original window.


```vb
Sub WinNum() 
 Dim lwindowNum As Long 
 
 lwindowNum = ActiveDocument.ActiveWindow.WindowNumber 
 NewWindow 
 ActiveDocument.Windows(lwindowNum).Activate 
End Sub
```


## See also


#### Concepts


[Window Object](window-object-word.md)


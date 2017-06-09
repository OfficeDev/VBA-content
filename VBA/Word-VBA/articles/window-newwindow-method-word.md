---
title: Window.NewWindow Method (Word)
keywords: vbawd10.chm157417577
f1_keywords:
- vbawd10.chm157417577
ms.prod: word
api_name:
- Word.Window.NewWindow
ms.assetid: f0a1b56a-2e6e-9849-24a2-2078424aa30a
ms.date: 06/08/2017
---


# Window.NewWindow Method (Word)

Opens a new window with the same document as the specified window. Returns a  **Window** object.


## Syntax

 _expression_ . **NewWindow**

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Return Value

Window


## Remarks

A colon (:) and a number appear in the window caption when more than one window is open for a the same document. The following two instructions are functionally equivalent.


```vb
Set myWindow = ActiveDocument.ActiveWindow.NewWindow 
Set myWindow = NewWindow
```


## Example

This example posts a message that indicates the number of windows that exist before and after you open a new window for Document1.


```vb
MsgBox Windows.Count &; " windows open" 
Windows("Document1").NewWindow 
MsgBox Windows.Count &; " windows open"
```


## See also


#### Concepts


[Window Object](window-object-word.md)


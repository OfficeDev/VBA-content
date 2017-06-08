---
title: Application.NewWindow Method (Word)
keywords: vbawd10.chm158335321
f1_keywords:
- vbawd10.chm158335321
ms.prod: word
api_name:
- Word.Application.NewWindow
ms.assetid: 0af15be1-7002-bd73-13da-19635d09b034
ms.date: 06/08/2017
---


# Application.NewWindow Method (Word)

Opens a new window with the same document as the specified window. Returns a  **Window** object.


## Syntax

 _expression_ . **NewWindow**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Return Value

Window


## Remarks

A colon (:) and a number appear in the window caption when more than one window is open for a document. If the  **NewWindow** method is used with the **Application** object, a new window is opened for the active window. The following two instructions are functionally equivalent.


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

This example opens a new window, arranges all the open windows, closes the new window, and then rearranges the open windows.




```vb
Set myWindow = NewWindow 
Windows.Arrange ArrangeStyle:=wdTiled 
myWindow.Close 
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


#### Concepts


[Application Object](application-object-word.md)


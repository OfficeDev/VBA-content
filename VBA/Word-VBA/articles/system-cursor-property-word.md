---
title: System.Cursor Property (Word)
keywords: vbawd10.chm154468368
f1_keywords:
- vbawd10.chm154468368
ms.prod: word
api_name:
- Word.System.Cursor
ms.assetid: f4acf757-920f-f389-948e-e2a142d451b0
ms.date: 06/08/2017
---


# System.Cursor Property (Word)

Returns or sets the state (shape) of the pointer. Can be one of the following  **WdCursorType** constants: **wdCursorIBeam** , **wdCursorNormal** , **wdCursorNorthwestArrow** , or **wdCursorWait** . Read/write **Long** .


## Syntax

 _expression_ . **Cursor**

 _expression_ A variable that represents a **[System](system-object-word.md)** object.


## Example

This example prints a message on the status bar and changes the pointer to a busy pointer.


```vb
Dim intWait As Integer 
 
StatusBar = "Please wait..." 
 
For intWait = 1 To 1000 
 System.Cursor = wdCursorWait 
Next intWait 
 
StatusBar = "Task completed" 
System.Cursor = wdCursorNormal
```


## See also


#### Concepts


[System Object](system-object-word.md)


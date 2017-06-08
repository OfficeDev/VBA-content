---
title: Dialog.CommandName Property (Word)
keywords: vbawd10.chm163085574
f1_keywords:
- vbawd10.chm163085574
ms.prod: word
api_name:
- Word.Dialog.CommandName
ms.assetid: 5bd7a993-b40e-57ca-65c7-260efcea488b
ms.date: 06/08/2017
---


# Dialog.CommandName Property (Word)

Returns the name of the procedure that displays the specified built-in dialog box. Read-only  **String** .


## Syntax

 _expression_ . **CommandName**

 _expression_ A variable that represents a **[Dialog](dialog-object-word.md)** object.


## Remarks

For more information about working with built-in Word dialog boxes, see [Displaying Built-in Word Dialog Boxes](http://msdn.microsoft.com/library/abe465f9-09a1-72ea-2e2d-9de14fc02434%28Office.15%29.aspx).


## Example

This example displays the name of the procedure that displays the  **Save As** dialog box ( **File** menu): **FileSaveAs** .


```vb
MsgBox Dialogs(wdDialogFileSaveAs).CommandName
```


## See also


#### Concepts


[Dialog Object](dialog-object-word.md)


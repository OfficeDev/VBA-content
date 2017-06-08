---
title: MailMerge.EditMainDocument Method (Word)
keywords: vbawd10.chm153092205
f1_keywords:
- vbawd10.chm153092205
ms.prod: word
api_name:
- Word.MailMerge.EditMainDocument
ms.assetid: 06ef9288-9434-7e75-ca6c-75c21fffd6b4
ms.date: 06/08/2017
---


# MailMerge.EditMainDocument Method (Word)

Activates the mail merge main document associated with the specified header source or data source document.


## Syntax

 _expression_ . **EditMainDocument**

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Remarks

If the main document isn't open, an error occurs. Use the  **Open** method if the main document isn't currently open.


## Example

This example attempts to activate the main document associated with the active data source document. If the main document isn't open, the  **Open** dialog box is displayed, with a message in the status bar.


```vb
Sub ActivateMain() 
 On Error GoTo errorhandler 
 Documents("Data.doc").MailMerge.EditMainDocument 
 
 Exit Sub 
 
errorhandler: 
 If Err = 4605 Then StatusBar = "Main document is not open" 
 Dialogs(wdDialogFileOpen).Show 
End Sub
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)


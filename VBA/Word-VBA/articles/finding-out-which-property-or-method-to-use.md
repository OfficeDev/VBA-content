---
title: Finding Out Which Property or Method to Use
ms.prod: word
ms.assetid: 6da49a9c-e28f-dae5-f4bd-3124004052fb
ms.date: 06/08/2017
---


# Finding Out Which Property or Method to Use

You can use the macro recorder to learn which methods or properties you need to accomplish a task in Word. The macro recorder is a tool that translates your actions into Visual Basic instructions. For example, if you turn on the macro recorder and open a document named "Examples.doc", the macro recorder records an instruction similar to the following.


```vb
Sub Macro1() 
' 
' Macro1 Macro 
' Macro recorded 9/22/2000 by JeffSmith 
' 
 Documents.Open FileName:="Examples.doc", ConfirmConversions:=False, _ 
 ReadOnly:=False, AddToRecentFiles:=False, _ 
 PasswordDocument:="", PasswordTemplate:="", _ 
 Revert:=False, WritePasswordDocument:="", _ 
 WritePasswordTemplate:="", Format:=wdOpenFormatAuto 
End Sub
```


The  **[Documents](application-documents-property-word.md)** property returns the **[Documents](documents-object-word.md)** collection and the **[Open](documents-open-method-word.md)** method opens the specified file name. When you are first learning Visual Basic, using the macro recorder can help you learn which properties and methods you need to use to accomplish a task.

For more information, see  [Revising recorded Visual Basic macros](revising-recorded-visual-basic-macros.md).


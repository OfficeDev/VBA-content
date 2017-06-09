---
title: Document.VBASigned Property (Word)
keywords: vbawd10.chm158007631
f1_keywords:
- vbawd10.chm158007631
ms.prod: word
api_name:
- Word.Document.VBASigned
ms.assetid: aa00c1ad-8c1e-5f47-de42-72db8292d5c0
ms.date: 06/08/2017
---


# Document.VBASigned Property (Word)

 **True** if the Microsoft Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only **Boolean** .


## Syntax

 _expression_ . **VBASigned**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example loads a document called "Temp.doc" and tests to see whether or not it has a digital signature. If there is no digital signature, the example displays a warning message.


```vb
Documents.Open FileName:="C:\My Documents\Temp.doc" 
If ActiveDocument.VBASigned = False Then 
 MsgBox "Warning! This document " _ 
 &; "has not been digitally signed.", _ 
 vbCritical, "Digital Signature Warning" 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)


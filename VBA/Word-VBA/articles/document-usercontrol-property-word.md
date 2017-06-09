---
title: Document.UserControl Property (Word)
keywords: vbawd10.chm158007388
f1_keywords:
- vbawd10.chm158007388
ms.prod: word
api_name:
- Word.Document.UserControl
ms.assetid: 34ab71eb-397e-4c14-dfbe-d3f29f84c753
ms.date: 06/08/2017
---


# Document.UserControl Property (Word)

 **True** if the document was created or opened by the user. Read/write **Boolean** .


## Syntax

 _expression_ . **UserControl**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

This property returns  **False** if the document was created or opened programmatically from another Microsoft Office application with the **Open** method or the Visual Basic **CreateObject** or **GetObject** command.


 **Note**  If Word is visible to the user or if you call the  **UserControl** property from within a Word code module, this property will always return **True** .


## Example

This example displays the status of the  **UserControl** property for the active document. This example will only work correctly when run from another Office application with the Word object library loaded.


```vb
Set wd = New Word.Application 
Set wdDoc = _ 
 wd.Documents.Open("C:\My Documents\doc1.doc") 
If wdDoc.UserControl = True Then 
 MsgBox "This document was created or opened by the user." 
Else 
 MsgBox "This document was created programmatically." 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)


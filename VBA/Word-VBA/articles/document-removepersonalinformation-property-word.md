---
title: Document.RemovePersonalInformation Property (Word)
keywords: vbawd10.chm158007640
f1_keywords:
- vbawd10.chm158007640
ms.prod: word
api_name:
- Word.Document.RemovePersonalInformation
ms.assetid: cea369d5-6ccd-8326-abdc-c834c5b17975
ms.date: 06/08/2017
---


# Document.RemovePersonalInformation Property (Word)

 **True** if Microsoft Word removes all user information from comments, revisions, and the Properties dialog box upon saving a document. Read/write **Boolean** .


## Syntax

 _expression_ . **RemovePersonalInformation**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the current document to remove personal information from the document the next time the user saves it.


```vb
Sub RemovePersonalInfo() 
 ActiveDocument.RemovePersonalInformation = True 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)


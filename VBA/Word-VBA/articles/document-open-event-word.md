---
title: Document.Open Event (Word)
keywords: vbawd10.chm4001005
f1_keywords:
- vbawd10.chm4001005
ms.prod: word
api_name:
- Word.Document.Open
ms.assetid: 80ad090c-69bf-b50e-3171-eab5414309a2
ms.date: 06/08/2017
---


# Document.Open Event (Word)

Occurs when a document is opened.


## Syntax

Private Sub  _expression_ _**Private Sub Document_Open**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If the event procedure is stored in a template, the procedure will run when a new document based on that template is opened and when the template itself is opened as a document.

For information about using events with the  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## Example

This example displays a message when a document is opened. (The procedure can be stored in the  **ThisDocument** class module of a document or its attached template.)


```vb
Private Sub Document_Open() 
 MsgBox "This document is copyrighted." 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)


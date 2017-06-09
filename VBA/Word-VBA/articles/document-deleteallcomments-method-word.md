---
title: Document.DeleteAllComments Method (Word)
keywords: vbawd10.chm158007667
f1_keywords:
- vbawd10.chm158007667
ms.prod: word
api_name:
- Word.Document.DeleteAllComments
ms.assetid: 8c0bf7fa-a4de-91e0-3e2b-bb5d8897534a
ms.date: 06/08/2017
---


# Document.DeleteAllComments Method (Word)

Deletes all comments from the  **[Comments](comments-object-word.md)** collection in a document.


## Syntax

 _expression_ . **DeleteAllComments**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **Add** method for the **[Comments](comments-object-word.md)** object to add a comment to a document.


## Example

This example deletes all comments in the active document. This example assumes you have comments in active document.


```vb
Sub DelAllComments() 
 ActiveDocument.DeleteAllComments 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)


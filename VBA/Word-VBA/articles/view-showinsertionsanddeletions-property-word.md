---
title: View.ShowInsertionsAndDeletions Property (Word)
keywords: vbawd10.chm161808420
f1_keywords:
- vbawd10.chm161808420
ms.prod: word
api_name:
- Word.View.ShowInsertionsAndDeletions
ms.assetid: 3738a713-819d-5dfd-a197-8c97a3de5ab4
ms.date: 06/08/2017
---


# View.ShowInsertionsAndDeletions Property (Word)

 **True** for Microsoft Word to display insertions and deletions that were made to a document with Track Changes enabled. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowInsertionsAndDeletions**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example hides the insertions and deletions made in a document. This example assumes that the document in the active window contains revisions made by one or more reviewers.


```vb
Sub HideInsertDelete() 
 ActiveWindow.View.ShowInsertionsAndDeletions = False 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)


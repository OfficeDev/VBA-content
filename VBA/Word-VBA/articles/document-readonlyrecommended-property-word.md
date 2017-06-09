---
title: Document.ReadOnlyRecommended Property (Word)
keywords: vbawd10.chm158007348
f1_keywords:
- vbawd10.chm158007348
ms.prod: word
api_name:
- Word.Document.ReadOnlyRecommended
ms.assetid: d7190307-c58a-fa7a-7bb0-56478eac8160
ms.date: 06/08/2017
---


# Document.ReadOnlyRecommended Property (Word)

 **True** if Microsoft Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only. Read/write **Boolean** .


## Syntax

 _expression_ . **ReadOnlyRecommended**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets Word to suggest, when it is opening the document, that the document be opened as read-only.


```vb
ActiveDocument.ReadOnlyRecommended = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)


---
title: Document.KerningByAlgorithm Property (Word)
keywords: vbawd10.chm158007605
f1_keywords:
- vbawd10.chm158007605
ms.prod: word
api_name:
- Word.Document.KerningByAlgorithm
ms.assetid: b49416b2-bdb7-2e13-8243-9eb24cc51a2f
ms.date: 06/08/2017
---


# Document.KerningByAlgorithm Property (Word)

 **True** if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write **Boolean** .


## Syntax

 _expression_ . **KerningByAlgorithm**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets Microsoft Word to kern half-width Latin characters and punctuation marks in the active document.


```vb
ActiveDocument.KerningByAlgorithm = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)


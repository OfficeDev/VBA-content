---
title: Document.CheckConsistency Method (Word)
keywords: vbawd10.chm158007555
f1_keywords:
- vbawd10.chm158007555
ms.prod: word
api_name:
- Word.Document.CheckConsistency
ms.assetid: 9ae5e917-0bd7-7c20-ca00-eea5a7e9dff7
ms.date: 06/08/2017
---


# Document.CheckConsistency Method (Word)

Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.


## Syntax

 _expression_ . **CheckConsistency**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example checks the consistency of Japanese characters in the active document.


```vb
ActiveDocument.CheckConsistency
```


## See also


#### Concepts


[Document Object](document-object-word.md)


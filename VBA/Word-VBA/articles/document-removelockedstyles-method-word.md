---
title: Document.RemoveLockedStyles Method (Word)
keywords: vbawd10.chm158007783
f1_keywords:
- vbawd10.chm158007783
ms.prod: word
api_name:
- Word.Document.RemoveLockedStyles
ms.assetid: 0c20a3c9-b4b3-e9a6-06d1-a9bf9b16dc07
ms.date: 06/08/2017
---


# Document.RemoveLockedStyles Method (Word)

Purges a document of locked styles when formatting restrictions have been applied in a document.


## Syntax

 _expression_ . **RemoveLockedStyles**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

The following example purges the locked styles in the active document.


```vb
ActiveDocument.RemoveLockedStyles
```


## See also


#### Concepts


[Document Object](document-object-word.md)


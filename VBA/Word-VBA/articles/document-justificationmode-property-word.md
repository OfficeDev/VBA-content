---
title: Document.JustificationMode Property (Word)
keywords: vbawd10.chm158007606
f1_keywords:
- vbawd10.chm158007606
ms.prod: word
api_name:
- Word.Document.JustificationMode
ms.assetid: 17d1a45f-eab7-b9f4-99d7-b5a12c7acc10
ms.date: 06/08/2017
---


# Document.JustificationMode Property (Word)

Returns or sets the character spacing adjustment for the specified document. Read/write  **[WdJustificationMode](wdjustificationmode-enumeration-word.md)** .


## Syntax

 _expression_ . **JustificationMode**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets Microsoft Word to compress only punctuation marks when adjusting character spacing.


```vb
ActiveDocument.JustificationMode = wdJustificationModeCompressKana
```


## See also


#### Concepts


[Document Object](document-object-word.md)


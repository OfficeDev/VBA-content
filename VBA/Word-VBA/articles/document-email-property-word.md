---
title: Document.Email Property (Word)
keywords: vbawd10.chm158007615
f1_keywords:
- vbawd10.chm158007615
ms.prod: word
api_name:
- Word.Document.Email
ms.assetid: dd4f6a41-3ee6-c1bf-3a2c-e00a342e0009
ms.date: 06/08/2017
---


# Document.Email Property (Word)

Returns an  **[Email](email-object-word.md)** object that contains all the e-mail-related properties of the current document. Read-only.


## Syntax

 _expression_ . **Email**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example returns the name of the style associated with the current e-mail author.


```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```


## See also


#### Concepts


[Document Object](document-object-word.md)


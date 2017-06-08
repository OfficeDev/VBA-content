---
title: Document.EnforceStyle Property (Word)
keywords: vbawd10.chm158007767
f1_keywords:
- vbawd10.chm158007767
ms.prod: word
api_name:
- Word.Document.EnforceStyle
ms.assetid: ce2249ca-bdb0-f2b7-e9fa-a759c4507a74
ms.date: 06/08/2017
---


# Document.EnforceStyle Property (Word)

Returns or sets a  **Boolean** that represents whether formatting restrictions are enforced in a protected document.


## Syntax

 _expression_ . **EnforceStyle**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

The following example turns on formatting restrictions in the active document.


```vb
ActiveDocument.EnforceStyle = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)


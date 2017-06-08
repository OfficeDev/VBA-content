---
title: Selection.Document Property (Word)
keywords: vbawd10.chm158663659
f1_keywords:
- vbawd10.chm158663659
ms.prod: word
api_name:
- Word.Selection.Document
ms.assetid: 03b4bfd7-8d4a-f069-0c28-41be2ead8614
ms.date: 06/08/2017
---


# Selection.Document Property (Word)

Returns a  **[Document](document-object-word.md)** object associated with the specified selection. Read-only.


## Syntax

 _expression_ . **Document**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example displays the document name and path for the selection.


```
Msgbox Selection.Document.FullName
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)


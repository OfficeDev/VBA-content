---
title: Options.DocumentViewDirection Property (Word)
keywords: vbawd10.chm162988432
f1_keywords:
- vbawd10.chm162988432
ms.prod: word
api_name:
- Word.Options.DocumentViewDirection
ms.assetid: 5f68af9c-edff-1b6b-e111-954e9e845e62
ms.date: 06/08/2017
---


# Options.DocumentViewDirection Property (Word)

Returns or sets the alignment and reading order for the entire document. Read/write  **WdDocumentViewDirection** .


## Syntax

 _expression_ . **DocumentViewDirection**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the alignment to right and the reading order to right-to-left for the entire document.


```
Options.DocumentViewDirection = wdDocumentViewRtl
```


## See also


#### Concepts


[Options Object](options-object-word.md)


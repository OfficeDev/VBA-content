---
title: Document.NoLineBreakAfter Property (Word)
keywords: vbawd10.chm158007609
f1_keywords:
- vbawd10.chm158007609
ms.prod: word
api_name:
- Word.Document.NoLineBreakAfter
ms.assetid: 287a9e9e-355e-3faf-d7fb-ee68bb0e6568
ms.date: 06/08/2017
---


# Document.NoLineBreakAfter Property (Word)

Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write  **String** .


## Syntax

 _expression_ . **NoLineBreakAfter**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets "$", "(", "[", "\", and "{" as the kinsoku characters after which Microsoft Word will not break a line in the active document.


```vb
ActiveDocument.NoLineBreakAfter = "$([\{"
```


## See also


#### Concepts


[Document Object](document-object-word.md)


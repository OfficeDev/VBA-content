---
title: Document.NoLineBreakBefore Property (Word)
keywords: vbawd10.chm158007608
f1_keywords:
- vbawd10.chm158007608
ms.prod: word
api_name:
- Word.Document.NoLineBreakBefore
ms.assetid: 03d4bb24-1941-5f12-f9e5-bccdda37fb33
ms.date: 06/08/2017
---


# Document.NoLineBreakBefore Property (Word)

Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write  **String** .


## Syntax

 _expression_ . **NoLineBreakBefore**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets "!", ")", and "]" as the kinsoku characters before which Microsoft Word will not break a line in the active document.


```vb
ActiveDocument.NoLineBreakBefore = "!)]"
```


## See also


#### Concepts


[Document Object](document-object-word.md)


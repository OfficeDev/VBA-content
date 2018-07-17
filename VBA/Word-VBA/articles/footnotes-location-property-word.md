---
title: Footnotes.Location Property (Word)
keywords: vbawd10.chm155320420
f1_keywords:
- vbawd10.chm155320420
ms.prod: word
api_name:
- Word.Footnotes.Location
ms.assetid: bdbfc0e2-2c39-a7fd-675e-0ff6d50f0053
ms.date: 06/08/2017
---


# Footnotes.Location Property (Word)

Returns or sets the position of all footnotes. Read/write  **WdFootnoteLocation** .


## Syntax

 _expression_ . **Location**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Example

This example positions footnotes at the bottom of each page.


```vb
ActiveDocument.Footnotes.Location = wdBottomOfPage
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)


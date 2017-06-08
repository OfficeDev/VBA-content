---
title: FootnoteOptions.Location Property (Word)
keywords: vbawd10.chm170131556
f1_keywords:
- vbawd10.chm170131556
ms.prod: word
api_name:
- Word.FootnoteOptions.Location
ms.assetid: 29300e96-150f-ea6c-14ce-602816b6907a
ms.date: 06/08/2017
---


# FootnoteOptions.Location Property (Word)

Returns or sets the position of all footnotes. Read/write  **WdFootnoteLocation** .


## Syntax

 _expression_ . **Location**

 _expression_ Required. A variable that represents a **[Footnote](footnote-object-word.md)** object.


## Example

This example positions footnotes at the bottom of each page.


```vb
ActiveDocument.Footnotes.Location = wdBottomOfPage
```


## See also


#### Concepts


[FootnoteOptions Object](footnoteoptions-object-word.md)


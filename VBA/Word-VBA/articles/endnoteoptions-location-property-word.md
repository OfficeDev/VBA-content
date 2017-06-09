---
title: EndnoteOptions.Location Property (Word)
keywords: vbawd10.chm23593060
f1_keywords:
- vbawd10.chm23593060
ms.prod: word
api_name:
- Word.EndnoteOptions.Location
ms.assetid: 3fd348a5-69cd-7319-898e-3f1a102fd644
ms.date: 06/08/2017
---


# EndnoteOptions.Location Property (Word)

Returns or sets the position of all endnotes. Read/write  **WdEndnoteLocation** .


## Syntax

 _expression_ . **Location**

 _expression_ Required. A variable that represents an **[EndnoteOptions](endnoteoptions-object-word.md)** collection.


## Example

This example positions all endnotes at the end of sections.


```vb
ActiveDocument.Endnotes.Location = wdEndOfSection
```


## See also


#### Concepts


[EndnoteOptions Object](endnoteoptions-object-word.md)


---
title: Endnotes.Location Property (Word)
keywords: vbawd10.chm155254884
f1_keywords:
- vbawd10.chm155254884
ms.prod: word
api_name:
- Word.Endnotes.Location
ms.assetid: 948dd801-4ae3-0063-0bfd-28ea141d0b69
ms.date: 06/08/2017
---


# Endnotes.Location Property (Word)

Returns or sets the position of all endnotes. Read/write  **[WdEndnoteLocation](wdendnotelocation-enumeration-word.md)** . .


## Syntax

 _expression_ . **Location**

 _expression_ An expression that represents a **[Endnotes](endnotes-object-word.md)** object.


## Example

This example positions all endnotes at the end of sections.


```vb
ActiveDocument.Endnotes.Location = wdEndOfSection
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)


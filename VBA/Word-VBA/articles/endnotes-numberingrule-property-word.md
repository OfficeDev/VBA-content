---
title: Endnotes.NumberingRule Property (Word)
keywords: vbawd10.chm155254887
f1_keywords:
- vbawd10.chm155254887
ms.prod: word
api_name:
- Word.Endnotes.NumberingRule
ms.assetid: 8f21cc55-b065-86fc-0bc5-d54e9f0e58ac
ms.date: 06/08/2017
---


# Endnotes.NumberingRule Property (Word)

Returns or sets the way endnotes are numbered after page breaks or section breaks. Read/write  **[WdNumberingRule](wdnumberingrule-enumeration-word.md)** .


## Syntax

 _expression_ . **NumberingRule**

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


## Example

This example restarts endnote numbering after each section break in the active document.


```vb
ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)


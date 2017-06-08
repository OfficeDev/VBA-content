---
title: Footnotes.ResetSeparator Method (Word)
keywords: vbawd10.chm155320327
f1_keywords:
- vbawd10.chm155320327
ms.prod: word
api_name:
- Word.Footnotes.ResetSeparator
ms.assetid: 252633ab-a9a1-6dbe-7821-5c7969175996
ms.date: 06/08/2017
---


# Footnotes.ResetSeparator Method (Word)

Resets the footnote separator to the default separator.


## Syntax

 _expression_ . **ResetSeparator**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Remarks

The default separator is a short horizontal line that separates document text from notes.


## Example

This example resets the footnote separator to the default separator line.


```vb
ActiveDocument.Footnotes.ResetSeparator
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)


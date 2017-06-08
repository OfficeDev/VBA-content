---
title: Footnotes.ResetContinuationSeparator Method (Word)
keywords: vbawd10.chm155320328
f1_keywords:
- vbawd10.chm155320328
ms.prod: word
api_name:
- Word.Footnotes.ResetContinuationSeparator
ms.assetid: edb1dae6-3e62-b625-0982-64dec3b654c9
ms.date: 06/08/2017
---


# Footnotes.ResetContinuationSeparator Method (Word)

Resets the footnote or endnote continuation separator to the default separator.


## Syntax

 _expression_ . **ResetContinuationSeparator**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Remarks

The default separator is a long horizontal line that separates document text from notes continued from the previous page.


## Example

This example resets the footnote continuation separator to the default separator line.


```vb
ActiveDocument.Footnotes.ResetContinuationSeparator
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)


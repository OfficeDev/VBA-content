---
title: Endnotes.ResetSeparator Method (Word)
keywords: vbawd10.chm155254791
f1_keywords:
- vbawd10.chm155254791
ms.prod: word
api_name:
- Word.Endnotes.ResetSeparator
ms.assetid: 9d525a4b-d3ed-5a31-9c07-1c19129cd171
ms.date: 06/08/2017
---


# Endnotes.ResetSeparator Method (Word)

Resets the endnote separator to the default separator.


## Syntax

 _expression_ . **ResetSeparator**

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


## Remarks

 The default separator is a short horizontal line that separates document text from notes.


## Example

This example resets the endnote separator for the notes in the document where the selection is located.


```
Selection.Endnotes.ResetSeparator
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)


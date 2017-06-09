---
title: Paragraphs.HalfWidthPunctuationOnTopOfLine Property (Word)
keywords: vbawd10.chm156762232
f1_keywords:
- vbawd10.chm156762232
ms.prod: word
api_name:
- Word.Paragraphs.HalfWidthPunctuationOnTopOfLine
ms.assetid: 015e38d9-b376-29df-06de-ec3d36c553ca
ms.date: 06/08/2017
---


# Paragraphs.HalfWidthPunctuationOnTopOfLine Property (Word)

 **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **HalfWidthPunctuationOnTopOfLine**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets Microsoft Word to change punctuation symbols at the beginning of a line to half-width characters for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs.HalfWidthPunctuationOnTopOfLine = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


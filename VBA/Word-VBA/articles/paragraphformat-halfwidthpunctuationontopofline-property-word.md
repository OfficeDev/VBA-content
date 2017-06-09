---
title: ParagraphFormat.HalfWidthPunctuationOnTopOfLine Property (Word)
keywords: vbawd10.chm156434552
f1_keywords:
- vbawd10.chm156434552
ms.prod: word
api_name:
- Word.ParagraphFormat.HalfWidthPunctuationOnTopOfLine
ms.assetid: f4b2a723-ec3c-d8bd-eb82-66423399c549
ms.date: 06/08/2017
---


# ParagraphFormat.HalfWidthPunctuationOnTopOfLine Property (Word)

 **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **HalfWidthPunctuationOnTopOfLine**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets Microsoft Word to change punctuation symbols at the beginning of a line to half-width characters for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).HalfWidthPunctuationOnTopOfLine = True
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)


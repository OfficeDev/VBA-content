---
title: Paragraph.HalfWidthPunctuationOnTopOfLine Property (Word)
keywords: vbawd10.chm156696696
f1_keywords:
- vbawd10.chm156696696
ms.prod: word
api_name:
- Word.Paragraph.HalfWidthPunctuationOnTopOfLine
ms.assetid: 596456b6-cb29-9e9f-27ea-e8ad84d252f9
ms.date: 06/08/2017
---


# Paragraph.HalfWidthPunctuationOnTopOfLine Property (Word)

 **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **HalfWidthPunctuationOnTopOfLine**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets Microsoft Word to change punctuation symbols at the beginning of a line to half-width characters for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).HalfWidthPunctuationOnTopOfLine = True
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)


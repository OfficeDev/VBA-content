---
title: Paragraphs.Format Property (Word)
keywords: vbawd10.chm156763214
f1_keywords:
- vbawd10.chm156763214
ms.prod: word
api_name:
- Word.Paragraphs.Format
ms.assetid: 7f087836-82ad-829e-5529-258ba4a3a9b1
ms.date: 06/08/2017
---


# Paragraphs.Format Property (Word)

Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the formatting of the specified paragraph or paragraphs.


## Syntax

 _expression_ . **Format**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

The following example left-aligns all the paragraphs in the active document.


```vb
ActiveDocument.Paragraphs.Format.Alignment = wdAlignParagraphLeft
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


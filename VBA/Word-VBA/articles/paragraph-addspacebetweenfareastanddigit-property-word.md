---
title: Paragraph.AddSpaceBetweenFarEastAndDigit Property (Word)
keywords: vbawd10.chm156696698
f1_keywords:
- vbawd10.chm156696698
ms.prod: word
api_name:
- Word.Paragraph.AddSpaceBetweenFarEastAndDigit
ms.assetid: b4841607-2cf1-7607-8aca-c0e187a1d2dd
ms.date: 06/08/2017
---


# Paragraph.AddSpaceBetweenFarEastAndDigit Property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **AddSpaceBetweenFarEastAndDigit**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese text and numbers for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndDigit = True
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)


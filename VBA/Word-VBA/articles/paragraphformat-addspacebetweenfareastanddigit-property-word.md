---
title: ParagraphFormat.AddSpaceBetweenFarEastAndDigit Property (Word)
keywords: vbawd10.chm156434554
f1_keywords:
- vbawd10.chm156434554
ms.prod: word
api_name:
- Word.ParagraphFormat.AddSpaceBetweenFarEastAndDigit
ms.assetid: 9792aa0e-bb31-463b-ef7c-99847f587c19
ms.date: 06/08/2017
---


# ParagraphFormat.AddSpaceBetweenFarEastAndDigit Property (Word)

 **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **AddSpaceBetweenFarEastAndDigit**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese text and numbers for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndDigit = True
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)


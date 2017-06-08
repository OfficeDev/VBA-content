---
title: Paragraphs.HangingPunctuation Property (Word)
keywords: vbawd10.chm156762231
f1_keywords:
- vbawd10.chm156762231
ms.prod: word
api_name:
- Word.Paragraphs.HangingPunctuation
ms.assetid: e3a4005a-7a70-59c7-40d6-4e7489144b09
ms.date: 06/08/2017
---


# Paragraphs.HangingPunctuation Property (Word)

 **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **HangingPunctuation**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example enables hanging punctuation for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs.HangingPunctuation = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


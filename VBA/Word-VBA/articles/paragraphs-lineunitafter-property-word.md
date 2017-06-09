---
title: Paragraphs.LineUnitAfter Property (Word)
keywords: vbawd10.chm156762242
f1_keywords:
- vbawd10.chm156762242
ms.prod: word
api_name:
- Word.Paragraphs.LineUnitAfter
ms.assetid: 6cb3c9cc-bd98-7732-06b1-4108a542601e
ms.date: 06/08/2017
---


# Paragraphs.LineUnitAfter Property (Word)

Returns or sets the amount of spacing (in gridlines) after the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **LineUnitAfter**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the spacing after all paragraphs in the active document to one gridline.


```vb
ActiveDocument.Paragraphs.LineUnitAfter = 1
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


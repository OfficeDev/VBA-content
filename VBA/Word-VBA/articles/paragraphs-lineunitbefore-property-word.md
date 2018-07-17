---
title: Paragraphs.LineUnitBefore Property (Word)
keywords: vbawd10.chm156762241
f1_keywords:
- vbawd10.chm156762241
ms.prod: word
api_name:
- Word.Paragraphs.LineUnitBefore
ms.assetid: 8db3f0e4-1f52-ce37-b685-e8ace269d1d5
ms.date: 06/08/2017
---


# Paragraphs.LineUnitBefore Property (Word)

Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **LineUnitBefore**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the spacing before all paragraphs in the active document to one gridline.


```vb
ActiveDocument.Paragraphs.LineUnitBefore = 1
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


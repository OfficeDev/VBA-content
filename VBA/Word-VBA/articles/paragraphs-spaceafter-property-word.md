---
title: Paragraphs.SpaceAfter Property (Word)
keywords: vbawd10.chm156762224
f1_keywords:
- vbawd10.chm156762224
ms.prod: word
api_name:
- Word.Paragraphs.SpaceAfter
ms.assetid: 78a75278-acca-a588-0fef-01511cf67a20
ms.date: 06/08/2017
---


# Paragraphs.SpaceAfter Property (Word)

Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceAfter**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the spacing after all paragraphs in the active document to 12 points.


```vb
ActiveDocument.Paragraphs.SpaceAfter = 12
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


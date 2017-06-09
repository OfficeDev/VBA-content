---
title: Paragraphs.RightIndent Property (Word)
keywords: vbawd10.chm156762218
f1_keywords:
- vbawd10.chm156762218
ms.prod: word
api_name:
- Word.Paragraphs.RightIndent
ms.assetid: da5f408c-add9-05a6-bd3d-cd507af48312
ms.date: 06/08/2017
---


# Paragraphs.RightIndent Property (Word)

Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **RightIndent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the right indent for all paragraphs in the active document to 1 inch from the right margin. The  **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


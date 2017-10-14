---
title: Paragraphs.Last Property (Word)
keywords: vbawd10.chm156762116
f1_keywords:
- vbawd10.chm156762116
ms.prod: word
api_name:
- Word.Paragraphs.Last
ms.assetid: 9d9d173d-7d4f-ff23-35be-e3aeed85cc3c
ms.date: 06/08/2017
---


# Paragraphs.Last Property (Word)

Returns a  **Paragraph** object that represents the last item in the collection of paragraphs.


## Syntax

 _expression_ . **Last**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example formats the last paragraph in the active document to be right-aligned.


```vb
ActiveDocument.Paragraphs.Last.Alignment = wdAlignParagraphRight
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


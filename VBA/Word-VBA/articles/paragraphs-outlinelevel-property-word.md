---
title: Paragraphs.OutlineLevel Property (Word)
keywords: vbawd10.chm156762314
f1_keywords:
- vbawd10.chm156762314
ms.prod: word
api_name:
- Word.Paragraphs.OutlineLevel
ms.assetid: ed44b494-84aa-3419-cc3f-69b330ec6aeb
ms.date: 06/08/2017
---


# Paragraphs.OutlineLevel Property (Word)

Returns or sets the outline level for the specified paragraphs. Read/write  **[WdOutlineLevel](wdoutlinelevel-enumeration-word.md)** .


## Syntax

 _expression_ . **OutlineLevel**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

If a paragraph has a heading style applied to it (Heading 1 through Heading 9), the outline level is the same as the heading style and cannot be changed. Outline levels are visible only in outline view or the document map pane.


## Example

This example returns the outline level of all paragraphs in the active document.


```
temp = ActiveDocument.Paragraphs.OutlineLevel
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


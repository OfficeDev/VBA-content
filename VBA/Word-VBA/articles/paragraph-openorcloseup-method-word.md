---
title: Paragraph.OpenOrCloseUp Method (Word)
keywords: vbawd10.chm156696879
f1_keywords:
- vbawd10.chm156696879
ms.prod: word
api_name:
- Word.Paragraph.OpenOrCloseUp
ms.assetid: ab5a657f-9a8f-a191-76ac-f16aaa2758ee
ms.date: 06/08/2017
---


# Paragraph.OpenOrCloseUp Method (Word)

Toggles the spacing before a paragraph.


## Syntax

 _expression_ . **OpenOrCloseUp**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

If spacing before the specified paragraphs is 0 (zero), this method sets spacing to 12 points. If spacing before the paragraphs is greater than 0 (zero), this method sets spacing to 0 (zero).


## Example

This example toggles the formatting of the first paragraph in the active document to either add 12 points of space before the paragraph or leave no space before it.


```vb
ActiveDocument.Paragraphs(1).OpenOrCloseUp
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)


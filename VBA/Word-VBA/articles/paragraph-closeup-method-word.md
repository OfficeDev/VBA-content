---
title: Paragraph.CloseUp Method (Word)
keywords: vbawd10.chm156696877
f1_keywords:
- vbawd10.chm156696877
ms.prod: word
api_name:
- Word.Paragraph.CloseUp
ms.assetid: eb244d95-8de9-6de3-730d-663f6149c973
ms.date: 06/08/2017
---


# Paragraph.CloseUp Method (Word)

Removes any spacing before the specified paragraph.


## Syntax

 _expression_ . **CloseUp**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs(1).CloseUp 
ActiveDocument.Paragraphs(1).SpaceBefore = 0
```


## Example

This example removes any space before the first paragraph in the selection.


```
Selection.Paragraphs(1).CloseUp
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)


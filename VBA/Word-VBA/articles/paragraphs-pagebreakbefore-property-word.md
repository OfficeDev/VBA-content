---
title: Paragraphs.PageBreakBefore Property (Word)
keywords: vbawd10.chm156762216
f1_keywords:
- vbawd10.chm156762216
ms.prod: word
api_name:
- Word.Paragraphs.PageBreakBefore
ms.assetid: 573ff2bc-e9df-8a6e-49eb-0773e578969d
ms.date: 06/08/2017
---


# Paragraphs.PageBreakBefore Property (Word)

 **True** if a page break is forced before the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .


## Syntax

 _expression_ . **PageBreakBefore**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example forces a page break before the first paragraph in the selection.


```vb
Selection.Paragraphs.PageBreakBefore = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


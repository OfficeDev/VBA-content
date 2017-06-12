---
title: Paragraphs.Hyphenation Property (Word)
keywords: vbawd10.chm156762225
f1_keywords:
- vbawd10.chm156762225
ms.prod: word
api_name:
- Word.Paragraphs.Hyphenation
ms.assetid: 0437a95c-719d-72ce-1de9-ce9d3fd166ff
ms.date: 06/08/2017
---


# Paragraphs.Hyphenation Property (Word)

 **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long** .


## Syntax

 _expression_ . **Hyphenation**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property can be  **True** , **False** or **wdUndefined** .


## Example

This example turns off automatic hyphenation for all paragraphs in the active document.


```vb
ActiveDocument.Paragraphs.Hyphenation = False
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


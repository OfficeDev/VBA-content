---
title: Paragraphs.Count Property (Word)
keywords: vbawd10.chm156762114
f1_keywords:
- vbawd10.chm156762114
ms.prod: word
api_name:
- Word.Paragraphs.Count
ms.assetid: 8e2844f2-1a09-63d9-a981-e39a32a87d2f
ms.date: 06/08/2017
---


# Paragraphs.Count Property (Word)

Returns a  **Long** that represents the number of paragraphs in the collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example displays the number of paragraphs in the active document.


```vb
MsgBox "The active document contains " &; _ 
 ActiveDocument.Paragraphs.Count &; " paragraphs."
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


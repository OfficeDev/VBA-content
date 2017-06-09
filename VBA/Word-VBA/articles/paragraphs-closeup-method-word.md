---
title: Paragraphs.CloseUp Method (Word)
keywords: vbawd10.chm156762413
f1_keywords:
- vbawd10.chm156762413
ms.prod: word
api_name:
- Word.Paragraphs.CloseUp
ms.assetid: 0fa0afb7-fbdf-ab26-1b49-312f526d69c6
ms.date: 06/08/2017
---


# Paragraphs.CloseUp Method (Word)

Removes any spacing before the specified paragraphs.


## Syntax

 _expression_ . **CloseUp**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs.CloseUp 
ActiveDocument.Paragraphs.SpaceBefore = 0
```


## Example

This example removes any space before the first paragraph in the selection.


```
Selection.Paragraphs.CloseUp
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


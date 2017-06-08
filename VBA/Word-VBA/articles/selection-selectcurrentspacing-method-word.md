---
title: Selection.SelectCurrentSpacing Method (Word)
keywords: vbawd10.chm158663175
f1_keywords:
- vbawd10.chm158663175
ms.prod: word
api_name:
- Word.Selection.SelectCurrentSpacing
ms.assetid: 1a49caa6-d261-e9d7-9d64-c564c30a7e29
ms.date: 06/08/2017
---


# Selection.SelectCurrentSpacing Method (Word)

Extends the selection forward until a paragraph with different line spacing is encountered.


## Syntax

 _expression_ . **SelectCurrentSpacing**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example selects all consecutive paragraphs that have the same line spacing and changes the line spacing to single spacing.


```vb
With Selection 
 .SelectCurrentSpacing 
 .ParagraphFormat.Space1 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)


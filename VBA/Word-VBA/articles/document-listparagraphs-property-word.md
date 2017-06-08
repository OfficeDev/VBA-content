---
title: Document.ListParagraphs Property (Word)
keywords: vbawd10.chm158007380
f1_keywords:
- vbawd10.chm158007380
ms.prod: word
api_name:
- Word.Document.ListParagraphs
ms.assetid: 6e34e592-e745-95cd-8ffc-cd25f75db956
ms.date: 06/08/2017
---


# Document.ListParagraphs Property (Word)

Returns a  **ListParagraphs** object that represents all the numbered paragraphs in a document. Read-only.


## Syntax

 _expression_ . **ListParagraphs**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](http://msdn.microsoft.com/library/8c0b84c0-582b-32f7-68e0-6383d0661e74%28Office.15%29.aspx).


## Example

This example adds a yellow background to each numbered or bulleted paragraph in the first document.


```vb
For Each numpar In Documents(1).ListParagraphs 
 numpar.Shading.BackgroundPatternColorIndex = wdYellow 
Next numpar
```


## See also


#### Concepts


[Document Object](document-object-word.md)


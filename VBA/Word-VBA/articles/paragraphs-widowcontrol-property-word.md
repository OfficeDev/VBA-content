---
title: Paragraphs.WidowControl Property (Word)
keywords: vbawd10.chm156762226
f1_keywords:
- vbawd10.chm156762226
ms.prod: word
api_name:
- Word.Paragraphs.WidowControl
ms.assetid: 0e28845c-d65e-8f4a-6a5c-729622d2a9ec
ms.date: 06/08/2017
---


# Paragraphs.WidowControl Property (Word)

 **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be **True** , **False** or **wdUndefined** . Read/write **Long** .


## Syntax

 _expression_ . **WidowControl**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example formats the paragraphs in the active document so that the first or last line in a paragraph can appear by itself at the top or bottom of a page.


```vb
ActiveDocument.Paragraphs.WidowControl = False
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)


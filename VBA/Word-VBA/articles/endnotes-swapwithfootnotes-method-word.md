---
title: Endnotes.SwapWithFootnotes Method (Word)
keywords: vbawd10.chm155254790
f1_keywords:
- vbawd10.chm155254790
ms.prod: word
api_name:
- Word.Endnotes.SwapWithFootnotes
ms.assetid: b95f65e3-16aa-1290-f47c-6cfe1c7849d7
ms.date: 06/08/2017
---


# Endnotes.SwapWithFootnotes Method (Word)

Converts all endnotes in a document to footnotes and vice versa.


## Syntax

 _expression_ . **SwapWithFootnotes**

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


## Remarks

To convert a range of endnotes to footnotes, use the  **[Convert](endnotes-convert-method-word.md)** method.


## Example

This example converts the endnotes in the active document to footnotes and converts the footnotes to endnotes.


```vb
ActiveDocument.Endnotes.SwapWithFootnotes
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)


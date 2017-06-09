---
title: Footnotes.Convert Method (Word)
keywords: vbawd10.chm155320325
f1_keywords:
- vbawd10.chm155320325
ms.prod: word
api_name:
- Word.Footnotes.Convert
ms.assetid: 6d703b30-b0a5-bf33-4ae8-c6574cceae99
ms.date: 06/08/2017
---


# Footnotes.Convert Method (Word)

Converts endnotes to footnotes, or vice versa.


## Syntax

 _expression_ . **Convert**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** object.


## Example

This example converts the footnotes in the selection to endnotes.


```vb
If Selection.Footnotes.Count > 0 Then Selection.Footnotes.Convert
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)


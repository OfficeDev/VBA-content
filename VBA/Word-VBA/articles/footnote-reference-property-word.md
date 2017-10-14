---
title: Footnote.Reference Property (Word)
keywords: vbawd10.chm155123717
f1_keywords:
- vbawd10.chm155123717
ms.prod: word
api_name:
- Word.Footnote.Reference
ms.assetid: c13dfad2-a103-8d91-0e55-86022a7857cd
ms.date: 06/08/2017
---


# Footnote.Reference Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents a footnote reference mark.


## Syntax

 _expression_ . **Reference**

 _expression_ Required. A variable that represents a **[Footnote](footnote-object-word.md)** object.


## Example

This example sets  _myRange_ to the first footnote reference mark in the active document and then copies the reference mark.


```vb
Set myRange = ActiveDocument.Footnotes(1).Reference 
myRange.Copy
```


## See also


#### Concepts


[Footnote Object](footnote-object-word.md)


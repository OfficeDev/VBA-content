---
title: Endnote.Reference Property (Word)
keywords: vbawd10.chm155058181
f1_keywords:
- vbawd10.chm155058181
ms.prod: word
api_name:
- Word.Endnote.Reference
ms.assetid: 7e7bb259-8203-445c-fa84-80f1c05505d9
ms.date: 06/08/2017
---


# Endnote.Reference Property (Word)

Returns a  **Range** object that represents an endnote reference mark.


## Syntax

 _expression_ . **Reference**

 _expression_ Required. A variable that represents an **[Endnote](endnote-object-word.md)** object.


## Example

This example sets  _myRange_ to the first endnote reference mark in the active document and then copies the reference mark.


```vb
Set myRange = ActiveDocument.Endnotes(1).Reference 
myRange.Copy
```


## See also


#### Concepts


[Endnote Object](endnote-object-word.md)


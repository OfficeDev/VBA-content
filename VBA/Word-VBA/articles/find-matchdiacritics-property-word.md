---
title: Find.MatchDiacritics Property (Word)
keywords: vbawd10.chm162529381
f1_keywords:
- vbawd10.chm162529381
ms.prod: word
api_name:
- Word.Find.MatchDiacritics
ms.assetid: db03ebc8-32d7-bdb4-e4fa-257045ecc48b
ms.date: 06/08/2017
---


# Find.MatchDiacritics Property (Word)

 **True** if find operations match text with matching diacritics in a right-to-left language document. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchDiacritics**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Example

This example sets the current find operation to match diacritics.


```vb
Selection.Find.MatchDiacritics = True
```


## See also


#### Concepts


[Find Object](find-object-word.md)


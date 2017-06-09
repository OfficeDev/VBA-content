---
title: Find.MatchAlefHamza Property (Word)
keywords: vbawd10.chm162529382
f1_keywords:
- vbawd10.chm162529382
ms.prod: word
api_name:
- Word.Find.MatchAlefHamza
ms.assetid: 1023d28a-d6b7-658a-0fb2-e2f9bd11b457
ms.date: 06/08/2017
---


# Find.MatchAlefHamza Property (Word)

 **True** if find operations match text with matching alef hamzas in an Arabic language document. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchAlefHamza**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Example

This example sets the current find operation to match alef hamzas.


```vb
Selection.Find.MatchAlefHamza = True
```


## See also


#### Concepts


[Find Object](find-object-word.md)


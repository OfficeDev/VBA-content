---
title: Options.MatchFuzzyDash Property (Word)
keywords: vbawd10.chm162988345
f1_keywords:
- vbawd10.chm162988345
ms.prod: word
api_name:
- Word.Options.MatchFuzzyDash
ms.assetid: 141535f8-158d-c20c-34cf-6ed19a2601b2
ms.date: 06/08/2017
---


# Options.MatchFuzzyDash Property (Word)

 **True** if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyDash**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between minus signs, long vowel sounds, and dashes during a search.


```vb
Options.MatchFuzzyDash = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


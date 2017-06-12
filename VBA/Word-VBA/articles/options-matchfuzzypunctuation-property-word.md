---
title: Options.MatchFuzzyPunctuation Property (Word)
keywords: vbawd10.chm162988357
f1_keywords:
- vbawd10.chm162988357
ms.prod: word
api_name:
- Word.Options.MatchFuzzyPunctuation
ms.assetid: ea4cb188-7fd1-c7e5-e520-3f0826dc3cdd
ms.date: 06/08/2017
---


# Options.MatchFuzzyPunctuation Property (Word)

 **True** if Microsoft Word ignores the distinction between types of punctuation marks during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyPunctuation**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between types of punctuation marks during a search


```vb
Options.MatchFuzzyPunctuation = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


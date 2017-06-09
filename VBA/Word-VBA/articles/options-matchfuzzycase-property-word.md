---
title: Options.MatchFuzzyCase Property (Word)
keywords: vbawd10.chm162988341
f1_keywords:
- vbawd10.chm162988341
ms.prod: word
api_name:
- Word.Options.MatchFuzzyCase
ms.assetid: 2fa5cf3f-01d5-c47a-cc99-ce9249ea59bf
ms.date: 06/08/2017
---


# Options.MatchFuzzyCase Property (Word)

 **True** if Microsoft Word ignores the distinction between uppercase and lowercase letters during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyCase**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between uppercase and lowercase letters during a search.


```vb
Options.MatchFuzzyCase = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


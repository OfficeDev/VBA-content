---
title: Options.MatchFuzzySmallKana Property (Word)
keywords: vbawd10.chm162988344
f1_keywords:
- vbawd10.chm162988344
ms.prod: word
api_name:
- Word.Options.MatchFuzzySmallKana
ms.assetid: 743fdfa1-01da-32ee-22cf-c30852f382bf
ms.date: 06/08/2017
---


# Options.MatchFuzzySmallKana Property (Word)

 **True** if Microsoft Word ignores the distinction between diphthongs and double consonants during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzySmallKana**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between diphthongs and double consonants during a search.


```vb
Options.MatchFuzzySmallKana = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


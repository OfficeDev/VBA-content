---
title: Options.MatchFuzzyDZ Property (Word)
keywords: vbawd10.chm162988350
f1_keywords:
- vbawd10.chm162988350
ms.prod: word
api_name:
- Word.Options.MatchFuzzyDZ
ms.assetid: 4594528b-3855-512d-9738-878ce68c4bf7
ms.date: 06/08/2017
---


# Options.MatchFuzzyDZ Property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Symbol](images/fe274_ZA06051753.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)" and between "
![Symbol](images/fe276_ZA06051755.gif)" and "
![Symbol](images/fe277_ZA06051756.gif)" during a search. Read/write  **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyDZ**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Symbol](images/fe274_ZA06051753.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)" and between "
![Symbol](images/fe276_ZA06051755.gif)" and "
![Symbol](images/fe277_ZA06051756.gif)" during a search.


```vb
Options.MatchFuzzyDZ = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


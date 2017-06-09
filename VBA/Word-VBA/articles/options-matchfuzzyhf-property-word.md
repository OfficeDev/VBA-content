---
title: Options.MatchFuzzyHF Property (Word)
keywords: vbawd10.chm162988353
f1_keywords:
- vbawd10.chm162988353
ms.prod: word
api_name:
- Word.Options.MatchFuzzyHF
ms.assetid: fc818d98-8cdc-2dfe-9898-d019a01b2077
ms.date: 06/08/2017
---


# Options.MatchFuzzyHF Property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Symbol](images/fe283_ZA06051762.gif)
![Symbol](images/fe284_ZA06051763.gif)" and "
![Symbol](images/fe238_ZA06051718.gif)
![Symbol](images/fe284_ZA06051763.gif)" and between "
![Symbol](images/fe285_ZA06051764.gif)
![Symbol](images/fe284_ZA06051763.gif)" and "
![Symbol](images/fe267_ZA06051746.gif)
![Symbol](images/fe284_ZA06051763.gif)" during a search. Read/write  **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyHF**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Symbol](images/fe283_ZA06051762.gif)
![Symbol](images/fe284_ZA06051763.gif)" and "
![Symbol](images/fe238_ZA06051718.gif)
![Symbol](images/fe284_ZA06051763.gif)" and between "
![Symbol](images/fe285_ZA06051764.gif)
![Symbol](images/fe284_ZA06051763.gif)" and "
![Symbol](images/fe267_ZA06051746.gif)
![Symbol](images/fe284_ZA06051763.gif)" during a search.


```vb
Options.MatchFuzzyHF = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


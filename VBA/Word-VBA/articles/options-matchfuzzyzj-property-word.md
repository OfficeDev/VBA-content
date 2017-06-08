---
title: Options.MatchFuzzyZJ Property (Word)
keywords: vbawd10.chm162988354
f1_keywords:
- vbawd10.chm162988354
ms.prod: word
api_name:
- Word.Options.MatchFuzzyZJ
ms.assetid: 8f722df0-9fa4-3207-9cad-694cac2d955a
ms.date: 06/08/2017
---


# Options.MatchFuzzyZJ Property (Word)

 **True** if Microsoft Word ignores the distinction between "
![](images/fe286_ZA06051765.gif)" and "
![](images/fe287_ZA06051766.gif)
![Symbol](images/fe209_ZA06051695.gif)" and between "
![Symbol](images/fe288_ZA06051767.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)
![Symbol](images/fe209_ZA06051695.gif)" during a search. Read/write  **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyZJ**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![](images/fe286_ZA06051765.gif)" and "
![](images/fe287_ZA06051766.gif)
![Symbol](images/fe209_ZA06051695.gif)" and between "
![Symbol](images/fe288_ZA06051767.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)
![Symbol](images/fe209_ZA06051695.gif)" during a search.


```vb
Options.MatchFuzzyZJ = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


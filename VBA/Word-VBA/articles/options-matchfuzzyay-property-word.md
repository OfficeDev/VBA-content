---
title: Options.MatchFuzzyAY Property (Word)
keywords: vbawd10.chm162988355
f1_keywords:
- vbawd10.chm162988355
ms.prod: word
api_name:
- Word.Options.MatchFuzzyAY
ms.assetid: f9a56522-f3a8-0527-e0e9-9144ccc468bc
ms.date: 06/08/2017
---


# Options.MatchFuzzyAY Property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Symbol](images/fe289_ZA06051768.gif)" and "
![Symbol](images/fe241_ZA06051721.gif)" following 
![Symbol](images/fe144_ZA06051649.gif)-row and 
![Symbol](images/fe209_ZA06051695.gif)-row characters during a search. Read/write  **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyAY**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Symbol](images/fe289_ZA06051768.gif)" and "
![Symbol](images/fe241_ZA06051721.gif)" following 
![Symbol](images/fe144_ZA06051649.gif)-row and 
![Symbol](images/fe209_ZA06051695.gif)-row characters during a search.


```vb
Options.MatchFuzzyAY = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


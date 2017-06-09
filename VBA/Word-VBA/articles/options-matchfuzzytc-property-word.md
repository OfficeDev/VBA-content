---
title: Options.MatchFuzzyTC Property (Word)
keywords: vbawd10.chm162988352
f1_keywords:
- vbawd10.chm162988352
ms.prod: word
api_name:
- Word.Options.MatchFuzzyTC
ms.assetid: 9dc9eb01-d530-f2ac-0bb7-27630ca3ad60
ms.date: 06/08/2017
---


# Options.MatchFuzzyTC Property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Symbol](images/fe279_ZA06051758.gif)
![Symbol](images/fe280_ZA06051759.gif)", "
![Symbol](images/fe281_ZA06051760.gif)
![Symbol](images/fe280_ZA06051759.gif)", and "
![Symbol](images/fe208_ZA06051694.gif)", and between "
![Symbol](images/fe282_ZA06051761.gif)
![Symbol](images/fe280_ZA06051759.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)" during a search. Read/write  **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyTC**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Symbol](images/fe279_ZA06051758.gif)
![Symbol](images/fe280_ZA06051759.gif)", "
![Symbol](images/fe281_ZA06051760.gif)
![Symbol](images/fe280_ZA06051759.gif)", and "
![Symbol](images/fe208_ZA06051694.gif)", and between "
![Symbol](images/fe282_ZA06051761.gif)
![Symbol](images/fe280_ZA06051759.gif)" and "
![Symbol](images/fe275_ZA06051754.gif)" during a search.


```vb
Options.MatchFuzzyTC = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


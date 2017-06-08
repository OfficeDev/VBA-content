---
title: Options.MatchFuzzyByte Property (Word)
keywords: vbawd10.chm162988342
f1_keywords:
- vbawd10.chm162988342
ms.prod: word
api_name:
- Word.Options.MatchFuzzyByte
ms.assetid: 978d49df-a417-11b8-069e-1147067cd1ed
ms.date: 06/08/2017
---


# Options.MatchFuzzyByte Property (Word)

 **True** if Microsoft Word ignores the distinction between full-width and half-width characters (Latin or Japanese) during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzyByte**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between full-width and half-width characters (Latin or Japanese) during a search.


```vb
Options.MatchFuzzyByte = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)


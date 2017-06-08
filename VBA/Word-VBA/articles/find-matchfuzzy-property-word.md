---
title: Find.MatchFuzzy Property (Word)
keywords: vbawd10.chm162529320
f1_keywords:
- vbawd10.chm162529320
ms.prod: word
api_name:
- Word.Find.MatchFuzzy
ms.assetid: 7f3e2fb7-1485-a945-7161-e4ccc62e25e8
ms.date: 06/08/2017
---


# Find.MatchFuzzy Property (Word)

 **True** if Microsoft Word uses the nonspecific search options for Japanese text during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchFuzzy**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Example

This example conducts a nonspecific search for "ピアノ" in the selected range and selects the next occurrence.


```vb
With Selection.Find 
    .ClearFormatting 
    .Text = "ピアノ" 
    .MatchFuzzy = True 
    .Execute Format:=False, Forward:=True, Wrap:=wdFindContinue 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)


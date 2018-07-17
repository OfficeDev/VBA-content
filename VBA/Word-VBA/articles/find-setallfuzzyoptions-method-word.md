---
title: Find.SetAllFuzzyOptions Method (Word)
keywords: vbawd10.chm162529312
f1_keywords:
- vbawd10.chm162529312
ms.prod: word
api_name:
- Word.Find.SetAllFuzzyOptions
ms.assetid: 3fb439eb-5f98-620e-0e16-5905a2b105c6
ms.date: 06/08/2017
---


# Find.SetAllFuzzyOptions Method (Word)

Activates all nonspecific search options associated with Japanese text.


## Syntax

 _expression_ . **SetAllFuzzyOptions**

 _expression_ Required. A variable that represents a **[Find](find-object-word.md)** object.


## Remarks

This method sets the following properties to  **True** :



| **MatchFuzzyAY**| **MatchFuzzyKanji**|
| **MatchFuzzyBV**| **MatchFuzzyKiKu**|
| **MatchFuzzyByte**| **MatchFuzzyOldKana**|
| **MatchFuzzyCase**| **MatchFuzzyProlongedSoundMark**|
| **MatchFuzzyDash**| **MatchFuzzyPunctuation**|
| **MatchFuzzyDZ**| **MatchFuzzySmallKana**|
| **MatchFuzzyHF**| **MatchFuzzySpace**|
| **MatchFuzzyHiragana**| **MatchFuzzyTC**|
| **MatchFuzzyIterationMark**| **MatchFuzzyZJ**|

## Example

This example activates all nonspecific options before executing a search in the selected range. If the word "バイオリン" is formatted as bold, the entire paragraph is selected and copied to the Clipboard.


```vb
With Selection.Find 
    .ClearFormatting 
    .SetAllFuzzyOptions 
    .Font.Bold = True 
    .Execute FindText:=" バイオリン", Format:=True, Forward:=True 
    If .Found = True Then 
        .Parent.Expand Unit:=wdParagraph 
        .Parent.Copy 
    End If 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)


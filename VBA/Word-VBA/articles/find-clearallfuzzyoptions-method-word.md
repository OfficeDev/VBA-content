---
title: Find.ClearAllFuzzyOptions Method (Word)
keywords: vbawd10.chm162529313
f1_keywords:
- vbawd10.chm162529313
ms.prod: word
api_name:
- Word.Find.ClearAllFuzzyOptions
ms.assetid: cf0b33a4-bfcc-36f9-e4b4-b98b3c628c0d
ms.date: 06/08/2017
---


# Find.ClearAllFuzzyOptions Method (Word)

Clears all nonspecific search options associated with Japanese text.


## Syntax

 _expression_ . **ClearAllFuzzyOptions**

 _expression_ Required. A variable that represents a **[Find](find-object-word.md)** object.


## Remarks

This method sets the following properties to  **False** :



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

This example clears all nonspecific search options before executing a search in the selected range. If the word "バイオリン" is formatted as bold, the entire paragraph will be selected and copied to the Clipboard.


```vb
With Selection.Find 
    .ClearFormatting 
    .ClearAllFuzzyOptions 
    .Font.Bold = True 
    .Execute FindText:="バイオリン", Format:=True, Forward:=True 
    If .Found = True Then 
        .Parent.Expand Unit:=wdParagraph 
        .Parent.Copy 
    End If 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)


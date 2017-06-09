---
title: Find.MatchByte Property (Word)
keywords: vbawd10.chm162529321
f1_keywords:
- vbawd10.chm162529321
ms.prod: word
api_name:
- Word.Find.MatchByte
ms.assetid: c7da111f-e3ea-dec9-8091-5ccd9cd63cc7
ms.date: 06/08/2017
---


# Find.MatchByte Property (Word)

 **True** if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchByte**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Example

This example searches for the term "マイクロソフト" in the specified range without distinguishing between full-width and half-width characters.


```vb
With Selection.Find 
    .ClearFormatting 
    .MatchWholeWord = True 
    .MatchByte = False 
    .Execute FindText:="マイクロソフト" 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)


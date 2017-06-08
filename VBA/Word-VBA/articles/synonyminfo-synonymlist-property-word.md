---
title: SynonymInfo.SynonymList Property (Word)
keywords: vbawd10.chm161153031
f1_keywords:
- vbawd10.chm161153031
ms.prod: word
api_name:
- Word.SynonymInfo.SynonymList
ms.assetid: c51a5a79-9724-531b-acca-7e8b6c3feff9
ms.date: 06/08/2017
---


# SynonymInfo.SynonymList Property (Word)

Returns a list of synonyms for a specified meaning of a word or phrase. The list is returned as an array of strings. Read-only  **Variant** .


## Syntax

 _expression_ . **SynonymList**( **_Meaning_** )

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Example

This example returns a list of synonyms for the word "big," using the meaning "generous" in U.S. English.


```vb
Slist = SynonymInfo(Word:="big", LanguageID:=wdEnglishUS) _ 
 .SynonymList(Meaning:="generous") 
For i = 1 To UBound(Slist) 
 Msgbox Slist(i) 
Next i
```

This example returns a list of synonyms for the second meaning of the selected word or phrase and displays these synonyms in the Immediate window of the Visual Basic editor. If there is no second meaning or if there are no synonyms, this is stated in a message box.




```vb
Set mySi = Selection.Range.SynonymInfo 
If mySi.MeaningCount >= 2 Then 
 synList = mySi.SynonymList(Meaning:=2) 
 For i = 1 To UBound(synList) 
 Debug.Print synList(i) 
 Next i 
Else 
 MsgBox "There is no second meaning for this word or phrase." 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)


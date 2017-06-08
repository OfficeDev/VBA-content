---
title: SynonymInfo.Found Property (Word)
keywords: vbawd10.chm161153026
f1_keywords:
- vbawd10.chm161153026
ms.prod: word
api_name:
- Word.SynonymInfo.Found
ms.assetid: a69e196b-4db1-fae7-172f-92f00264443b
ms.date: 06/08/2017
---


# SynonymInfo.Found Property (Word)

 **True** if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase. Read-only **Boolean** .


## Syntax

 _expression_ . **Found**

 _expression_ A variable that represents a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Example

This example checks to see whether the thesaurus contains any synonym suggestions for the word "authorize."


```vb
Dim siTemp As SynonymInfo 
 
Set siTemp = SynonymInfo(Word:="authorize", _ 
 LanguageID:=wdEnglishUS) 
If siTemp.Found = True Then 
 Msgbox "The thesaurus has suggestions." 
Else 
 Msgbox "The thesaurus has no suggestions." 
End If
```

This example checks to see whether the thesaurus contains any synonym suggestions for the selection. If it does, the example displays the Thesaurus dialog box with the synonyms listed.




```vb
Dim siTemp As SynonymInfo 
 
Set siTemp = Selection.Range.SynonymInfo 
If siTemp.Found = True Then 
 Selection.Range.CheckSynonyms 
Else 
 Msgbox "The thesaurus has no suggestions." 
End If
```

This example removes formatting from the find criteria before searching the selection. If the word "Hello" with bold formatting is found, the entire paragraph is selected and copied to the Clipboard.




```vb
With Selection.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Execute FindText:="Hello", Format:=True, Forward:=True 
 If .Found = True Then 
 .Parent.Expand Unit:=wdParagraph 
 .Parent.Copy 
 End If 
End With
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)


---
title: SynonymInfo.MeaningCount Property (Word)
keywords: vbawd10.chm161153027
f1_keywords:
- vbawd10.chm161153027
ms.prod: word
api_name:
- Word.SynonymInfo.MeaningCount
ms.assetid: 8b4913e2-eed1-f69c-470b-1f6b3ea8c192
ms.date: 06/08/2017
---


# SynonymInfo.MeaningCount Property (Word)

Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns 0 (zero) if no meanings were found. Read-only  **Long** .


## Syntax

 _expression_ . **MeaningCount**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

Each meaning represents a unique list of synonyms for the word or phrase.

The lists of related words, related expressions, and antonyms aren't counted as entries in the list of meanings.


## Example

This example checks to see whether any meanings were found for the selection. If any were found, the list of meanings is displayed in the Immediate window of the Visual Basic Editor.


```vb
Set mySynInfo = Selection.Range.SynonymInfo 
If mySynInfo.MeaningCount <> 0 Then 
 myList = mySynInfo.MeaningList 
 For i = 1 To Ubound(myList) 
 Debug.Print myList(i) 
 Next i 
Else 
 Msgbox "There were no meanings found." 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)


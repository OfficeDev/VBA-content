---
title: SynonymInfo.MeaningList Property (Word)
keywords: vbawd10.chm161153028
f1_keywords:
- vbawd10.chm161153028
ms.prod: word
api_name:
- Word.SynonymInfo.MeaningList
ms.assetid: 43eec397-41e6-7b13-f267-ae3b4914ec02
ms.date: 06/08/2017
---


# SynonymInfo.MeaningList Property (Word)

Returns the list of meanings for the word or phrase. The list is returned as an array of strings. Read-only  **Variant** .


## Syntax

 _expression_ . **MeaningList**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

The lists of related words, related expressions, and antonyms aren't counted as entries in the list of meanings.


## Example

This example checks to see whether any meanings were found for the third word in MyDoc.doc. If so, the meanings are displayed in a series of message boxes.


```vb
Set mySyn = Documents("MyDoc.doc").Words(3).SynonymInfo 
If mySyn.MeaningCount <> 0 Then 
 myList = mySyn.MeaningList 
 For i = 1 To UBound(myList) 
 Msgbox myList(i) 
 Next i 
Else 
 Msgbox "There were no meanings found." 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)


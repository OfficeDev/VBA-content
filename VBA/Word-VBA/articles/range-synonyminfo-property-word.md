---
title: Range.SynonymInfo Property (Word)
keywords: vbawd10.chm157155483
f1_keywords:
- vbawd10.chm157155483
ms.prod: word
api_name:
- Word.Range.SynonymInfo
ms.assetid: b63d2a0b-baa1-306d-10ee-72223099a9f2
ms.date: 06/08/2017
---


# Range.SynonymInfo Property (Word)

Returns a  **SynonymInfo** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the contents of a range.


## Syntax

 _expression_ . **SynonymInfo**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example returns a list of synonyms for the selection's first meaning.


```vb
Slist = Selection.Range.SynonymInfo.SynonymList(Meaning:=1) 
For i = 1 To UBound(Slist) 
 Msgbox Slist(i) 
Next i
```


## See also


#### Concepts


[Range Object](range-object-word.md)


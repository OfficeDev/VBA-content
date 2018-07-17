---
title: Application.SynonymInfo Property (Word)
keywords: vbawd10.chm158335035
f1_keywords:
- vbawd10.chm158335035
ms.prod: word
api_name:
- Word.Application.SynonymInfo
ms.assetid: 7aff62c5-d962-23b5-0e86-ae566f72914c
ms.date: 06/08/2017
---


# Application.SynonymInfo Property (Word)

Returns a  **[SynonymInfo](synonyminfo-object-word.md)** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.


## Syntax

 _expression_ . **SynonymInfo**( **_Word_** , **_LanguageID_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**|The specified word or phrase.|
| _LanguageID_|Optional| **Variant**|The language used for the thesaurus. Can be one of the  **WdLanguageID** constants (although some of the constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed).|

## Example

This example returns a list of antonyms for the word "big" in U.S. English.


```vb
Alist = SynonymInfo(Word:="big", _ 
 LanguageID:=wdEnglishUS).AntonymList 
For i = 1 To UBound(Alist) 
 Msgbox Alist(i) 
Next i
```


## See also


#### Concepts


[Application Object](application-object-word.md)


---
title: Global.SynonymInfo Property (Word)
keywords: vbawd10.chm163119163
f1_keywords:
- vbawd10.chm163119163
ms.prod: word
api_name:
- Word.Global.SynonymInfo
ms.assetid: 792a9d40-2b03-6f3d-ed5e-2fc388a3b3d2
ms.date: 06/08/2017
---


# Global.SynonymInfo Property (Word)

Returns a  **SynonymInfo** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.


## Syntax

 _expression_ . **SynonymInfo**( **_Word_** , **_LanguageID_** )

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**|The specified word or phrase.|
| _LanguageID_|Optional| **Variant**|The language used for the thesaurus. Can be one of the  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constants (although some of the constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed).|

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


[Global Object](global-object-word.md)


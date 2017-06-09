---
title: Language.SpellingDictionaryType Property (Word)
keywords: vbawd10.chm158138387
f1_keywords:
- vbawd10.chm158138387
ms.prod: word
api_name:
- Word.Language.SpellingDictionaryType
ms.assetid: 4bde19be-a568-7145-f094-d483dc997020
ms.date: 06/08/2017
---


# Language.SpellingDictionaryType Property (Word)

Returns or sets the proofing tool type. Read/write  **WdDictionaryType** .


## Syntax

 _expression_ . **SpellingDictionaryType**

 _expression_ Required. A variable that represents a **[Language](language-object-word.md)** object.


## Remarks

You can use this property to change the active spelling dictionary to one of the available add-on dictionaries that work with Word. For example, there are legal, medical, and complete spelling dictionaries you can use instead of the standard dictionary.

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example returns the type of spelling dictionary used for U.S. English.


```
myType = Languages(wdEnglishUS).SpellingDictionaryType
```

This example makes the legal dictionary the active spelling dictionary.




```
Languages(wdEnglishUS).SpellingDictionaryType = wdSpellingLegal
```


## See also


#### Concepts


[Language Object](language-object-word.md)


---
title: AutoCorrect.HangulAndAlphabetAutoAdd Property (Word)
keywords: vbawd10.chm155779086
f1_keywords:
- vbawd10.chm155779086
ms.prod: word
api_name:
- Word.AutoCorrect.HangulAndAlphabetAutoAdd
ms.assetid: dbb1f1b7-21be-423a-e375-543c0c689034
ms.date: 06/08/2017
---


# AutoCorrect.HangulAndAlphabetAutoAdd Property (Word)

 **True** if Microsoft Word automatically adds words to the list of Hangul and alphabet AutoCorrect exceptions. Read/write **Boolean** .


## Syntax

 _expression_ . **HangulAndAlphabetAutoAdd**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

The list of Hangul and alphabet AutoCorrect exceptions is located on the  **Korean** tab in the **AutoCorrect Exceptions** dialog box. Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct.


## Example

This example sets Microsoft Word to automatically add words to the list of Hangul and alphabet AutoCorrect exceptions.


```vb
AutoCorrect.HangulAndAlphabetAutoAdd = True
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)


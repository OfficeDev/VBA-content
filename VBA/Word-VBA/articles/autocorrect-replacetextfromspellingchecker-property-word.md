---
title: AutoCorrect.ReplaceTextFromSpellingChecker Property (Word)
keywords: vbawd10.chm155779087
f1_keywords:
- vbawd10.chm155779087
ms.prod: word
api_name:
- Word.AutoCorrect.ReplaceTextFromSpellingChecker
ms.assetid: 8cc4a48f-86a6-5b26-ad2d-cca3b969047c
ms.date: 06/08/2017
---


# AutoCorrect.ReplaceTextFromSpellingChecker Property (Word)

 **True** if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types. Word only replaces words that contain a single misspelling and for which the spelling dictionary only lists one alternative. Read/write **Boolean** .


## Syntax

 _expression_ . **ReplaceTextFromSpellingChecker**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example sets Word to automatically replace misspelled text with suggestions from the spelling checker.


```vb
AutoCorrect.ReplaceTextFromSpellingChecker = True
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)


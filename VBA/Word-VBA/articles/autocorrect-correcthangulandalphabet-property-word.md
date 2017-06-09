---
title: AutoCorrect.CorrectHangulAndAlphabet Property (Word)
keywords: vbawd10.chm155779084
f1_keywords:
- vbawd10.chm155779084
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectHangulAndAlphabet
ms.assetid: b6dc4a8e-9245-0c29-370f-c6fcbb3a924a
ms.date: 06/08/2017
---


# AutoCorrect.CorrectHangulAndAlphabet Property (Word)

 **True** if Microsoft Word automatically applies the correct font to Latin words typed in the middle of Hangul text or vice versa. Read/write **Boolean** .


## Syntax

 _expression_ . **CorrectHangulAndAlphabet**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically apply the correct font to Latin words typed in the middle of Hangul text or vice versa.


```vb
AutoCorrect.CorrectHangulAndAlphabet = True
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)


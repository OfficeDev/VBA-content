---
title: AutoCorrect.CorrectInitialCaps Property (Word)
keywords: vbawd10.chm155779074
f1_keywords:
- vbawd10.chm155779074
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectInitialCaps
ms.assetid: 5f24b0a7-8b5a-3688-7dbf-7e7ad7adec3b
ms.date: 06/08/2017
---


# AutoCorrect.CorrectInitialCaps Property (Word)

 **True** if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, "WOrd" is corrected to "Word." Read/write **Boolean** .


## Syntax

 _expression_ . **CorrectInitialCaps**

 _expression_ A variable that represents an **[AutoCorrect](autocorrect-object-word.md)** object.


## Example

This example sets Word to automatically correct errors in initial capitalization.


```vb
AutoCorrect.CorrectInitialCaps = True
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)


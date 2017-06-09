---
title: AutoCorrect.FirstLetterAutoAdd Property (Word)
keywords: vbawd10.chm155779080
f1_keywords:
- vbawd10.chm155779080
ms.prod: word
api_name:
- Word.AutoCorrect.FirstLetterAutoAdd
ms.assetid: 17f51d86-405a-7188-eb8c-bfde5bdb386c
ms.date: 06/08/2017
---


# AutoCorrect.FirstLetterAutoAdd Property (Word)

 **True** if Word automatically adds abbreviations to the list of AutoCorrect First Letter exceptions. Read/write **Boolean** .


## Syntax

 _expression_ . **FirstLetterAutoAdd**

 _expression_ A variable that represents an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

Word adds an abbreviation to this list if you delete and then retype the letter that Word capitalized immediately after the period following the abbreviation.


## Example

This example prevents Word from automatically adding abbreviations to the list of AutoCorrect First Letter exceptions.


```vb
AutoCorrect.FirstLetterAutoAdd = False
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

